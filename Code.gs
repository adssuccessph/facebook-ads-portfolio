const SHEETS = {
  DEMOS: 'Demos',
  PACKAGES: 'Packages',
  LEADS: 'Leads',
  META_ADS: 'MetaAdsRequests',
  REVIEWS: 'Reviews',
  SETTINGS: 'Settings'
};

const DEMO_HEADERS = ['id', 'title', 'category', 'description', 'imageUrl', 'demoUrl', 'createdAt', 'updatedAt'];
const PACKAGE_HEADERS = ['id', 'category', 'name', 'price', 'description', 'inclusions', 'badge', 'ctaLabel', 'coverImageUrl', 'sampleImageUrls', 'active', 'featured', 'sortOrder', 'createdAt', 'updatedAt'];
const LEAD_HEADERS = ['id', 'createdAt', 'name', 'business', 'contact', 'service', 'budget', 'message', 'source', 'status', 'package', 'privacyConsent'];
const META_ADS_HEADERS = ['id', 'createdAt', 'name', 'email', 'contact', 'business', 'facebookPageUrl', 'websiteUrl', 'businessPortfolioStatus', 'adsAccountStatus', 'previousAds', 'objective', 'productService', 'offer', 'targetLocations', 'targetAudience', 'dailyBudget', 'runDuration', 'creativeStatus', 'referenceLink', 'notes', 'package', 'source', 'consent', 'privacyConsent', 'status', 'updatedAt'];
const REVIEW_HEADERS = ['id', 'createdAt', 'name', 'email', 'business', 'service', 'rating', 'review', 'privacyConsent', 'status', 'updatedAt'];
const SETTINGS_HEADERS = ['key', 'value'];

const UPLOAD_FOLDER_NAME = 'Ads Success PH Website Uploads';
const MAX_UPLOAD_BYTES = 4 * 1024 * 1024;
const PACKAGE_CATEGORIES = ['Website', 'Creative Images', 'Meta Ads'];
const PUBLIC_WRITE_LOCK_TIMEOUT_MS = 60000;
const PUBLIC_COOLDOWN_SECONDS = { lead: 30, metaAds: 60, review: 60 };
const PUBLIC_DUPLICATE_TTL_SECONDS = { lead: 120, metaAds: 300, review: 600 };

function setupApp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SHEETS.DEMOS, DEMO_HEADERS);
  ensureSheet_(ss, SHEETS.PACKAGES, PACKAGE_HEADERS);
  ensureSheet_(ss, SHEETS.LEADS, LEAD_HEADERS);
  ensureSheet_(ss, SHEETS.META_ADS, META_ADS_HEADERS);
  ensureSheet_(ss, SHEETS.REVIEWS, REVIEW_HEADERS);
  ensureSheet_(ss, SHEETS.SETTINGS, SETTINGS_HEADERS);
  getOrCreateUploadFolder_();

  const settings = ss.getSheetByName(SHEETS.SETTINGS);
  if (settings.getLastRow() === 1) {
    const defaults = [
      ['brandName', 'Ads Success PH'],
      ['logoUrl', ''],
      ['heroTitle', 'Digital services built to help small businesses get noticed.'],
      ['heroSubtitle', 'Custom websites, AI-assisted advertising creatives, product posting designs, and Meta Ads services—planned around your business goals and budget.'],
      ['messengerUrl', '']
    ];
    settings.getRange(2, 1, defaults.length, 2).setValues(defaults);
  }

  Logger.log('V6 setup complete. Privacy consent, anti-spam, duplicate protection, write locking, validation, and header protection are ready.');
}

function configureAdminPassword() {
  const PASSWORD = 'CHANGE_THIS_PASSWORD';
  if (PASSWORD === 'CHANGE_THIS_PASSWORD') {
    throw new Error('Edit configureAdminPassword() and replace CHANGE_THIS_PASSWORD before running.');
  }
  PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD_HASH', hash_(PASSWORD));
  Logger.log('Admin password configured.');
}

function doGet(e) {
  try {
    const action = String((e.parameter && e.parameter.action) || '');
    let result;

    switch (action) {
      case 'getPublicData':
        result = getPublicData_();
        break;
      case 'getAdminData':
        requireAdminToken_(e.parameter.token);
        result = getAdminData_();
        break;
      default:
        throw new Error('Unknown GET action.');
    }

    return json_({ ok: true, ...result });
  } catch (error) {
    return json_({ ok: false, error: error.message });
  }
}

function doPost(e) {
  try {
    const action = String((e.parameter && e.parameter.action) || '');
    const payload = parsePayload_(e);
    let result = {};

    switch (action) {
      case 'submitLead':
        result = submitLead_(payload);
        break;
      case 'submitMetaAdsRequest':
        result = submitMetaAdsRequest_(payload);
        break;
      case 'submitReview':
        result = submitReview_(payload);
        break;
      case 'adminLogin':
        result = adminLogin_(payload);
        break;
      case 'uploadDemoImage':
        requireAdminToken_(payload.token);
        result = uploadDemoImage_(payload);
        break;
      case 'uploadPackageImage':
        requireAdminToken_(payload.token);
        result = uploadPackageImage_(payload);
        break;
      case 'saveDemo':
        requireAdminToken_(payload.token);
        result = saveDemo_(payload);
        break;
      case 'deleteDemo':
        requireAdminToken_(payload.token);
        result = deleteDemo_(payload);
        break;
      case 'savePackage':
        requireAdminToken_(payload.token);
        result = savePackage_(payload);
        break;
      case 'deletePackage':
        requireAdminToken_(payload.token);
        result = deletePackage_(payload);
        break;
      case 'updateLeadStatus':
        requireAdminToken_(payload.token);
        result = updateLeadStatus_(payload);
        break;
      case 'updateMetaAdsStatus':
        requireAdminToken_(payload.token);
        result = updateMetaAdsStatus_(payload);
        break;
      case 'updateReviewStatus':
        requireAdminToken_(payload.token);
        result = updateReviewStatus_(payload);
        break;
      case 'deleteReview':
        requireAdminToken_(payload.token);
        result = deleteReview_(payload);
        break;
      case 'saveSettings':
        requireAdminToken_(payload.token);
        result = saveSettings_(payload);
        break;
      default:
        throw new Error('Unknown POST action.');
    }

    return json_({ ok: true, ...result });
  } catch (error) {
    return json_({ ok: false, error: error.message });
  }
}

function getPublicData_() {
  const approvedReviews = getRowsAsObjects_(SHEETS.REVIEWS)
    .filter(item => item.status === 'Approved')
    .reverse()
    .slice(0, 12)
    .map(item => ({
      id: item.id,
      createdAt: item.createdAt,
      name: item.name,
      business: item.business,
      service: item.service,
      rating: item.rating,
      review: item.review
    }));

  const activePackages = getRowsAsObjects_(SHEETS.PACKAGES)
    .filter(item => item.active === 'Yes')
    .sort((a, b) => {
      const categoryDiff = PACKAGE_CATEGORIES.indexOf(a.category) - PACKAGE_CATEGORIES.indexOf(b.category);
      if (categoryDiff !== 0) return categoryDiff;
      const orderDiff = (Number(a.sortOrder) || 0) - (Number(b.sortOrder) || 0);
      if (orderDiff !== 0) return orderDiff;
      return String(a.createdAt).localeCompare(String(b.createdAt));
    });

  return {
    demos: getRowsAsObjects_(SHEETS.DEMOS).reverse(),
    packages: activePackages,
    reviews: approvedReviews,
    settings: getSettingsObject_()
  };
}

function getAdminData_() {
  return {
    demos: getRowsAsObjects_(SHEETS.DEMOS).reverse(),
    packages: getRowsAsObjects_(SHEETS.PACKAGES)
      .sort((a, b) => {
        const categoryDiff = PACKAGE_CATEGORIES.indexOf(a.category) - PACKAGE_CATEGORIES.indexOf(b.category);
        if (categoryDiff !== 0) return categoryDiff;
        return (Number(a.sortOrder) || 0) - (Number(b.sortOrder) || 0);
      }),
    leads: getRowsAsObjects_(SHEETS.LEADS).reverse(),
    metaAdsRequests: getRowsAsObjects_(SHEETS.META_ADS).reverse(),
    reviews: getRowsAsObjects_(SHEETS.REVIEWS).reverse(),
    settings: getSettingsObject_()
  };
}

function submitLead_(payload) {
  if (clean_(payload.companyWebsite)) return { message: 'Inquiry received.' };

  const required = ['name', 'contact', 'service', 'message', 'privacyConsent'];
  required.forEach(key => {
    if (!String(payload[key] || '').trim()) throw new Error(`Missing required field: ${key}`);
  });

  if (payload.privacyConsent !== 'yes') throw new Error('Please agree to the Privacy Notice before submitting.');

  const name = clean_(payload.name);
  const business = clean_(payload.business);
  const contact = clean_(payload.contact);
  const service = clean_(payload.service);
  const budget = clean_(payload.budget);
  const message = clean_(payload.message);
  const packageName = clean_(payload.package);

  validateLength_(name, 'Name', 2, 100);
  validateLength_(business, 'Business name', 0, 150);
  validateLength_(contact, 'Contact', 3, 150);
  validateLength_(service, 'Service', 2, 150);
  validateLength_(budget, 'Budget', 0, 100);
  validateLength_(message, 'Project details', 10, 2000);
  validateLength_(packageName, 'Package', 0, 250);

  return withPublicWriteLock_(() => {
    enforceSubmissionProtection_('lead', contact, [name, business, contact, service, message]);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LEADS);
    appendObjectRow_(sheet, LEAD_HEADERS, {
      id: Utilities.getUuid(),
      createdAt: formatDate_(new Date()),
      name,
      business,
      contact,
      service,
      budget,
      message,
      source: clean_(payload.source || 'Website').slice(0, 150),
      status: 'New',
      package: packageName,
      privacyConsent: 'Yes'
    });

    markSubmissionProtection_('lead', contact, [name, business, contact, service, message]);
    return { message: 'Inquiry saved.' };
  });
}


function submitMetaAdsRequest_(payload) {
  if (clean_(payload.companyWebsite)) return { message: 'Meta Ads request received.' };

  const required = [
    'name', 'email', 'contact', 'business', 'businessPortfolioStatus',
    'adsAccountStatus', 'previousAds', 'objective', 'productService', 'offer',
    'targetLocations', 'targetAudience', 'dailyBudget', 'runDuration',
    'creativeStatus', 'consent', 'privacyConsent'
  ];

  required.forEach(key => {
    if (!String(payload[key] || '').trim()) throw new Error(`Missing required field: ${key}`);
  });

  const email = clean_(payload.email).toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Please enter a valid email address.');
  if (payload.consent !== 'yes') throw new Error('Please confirm the Meta Ads service and access acknowledgement.');
  if (payload.privacyConsent !== 'yes') throw new Error('Please agree to the Privacy Notice before submitting.');

  const allowedPortfolio = ['Already set up', 'Not set up yet', 'Not sure what this is'];
  const allowedAdAccounts = ['I already have an ad account', 'I do not have an ad account yet', 'Not sure'];
  if (!allowedPortfolio.includes(clean_(payload.businessPortfolioStatus))) throw new Error('Invalid Business Portfolio status.');
  if (!allowedAdAccounts.includes(clean_(payload.adsAccountStatus))) throw new Error('Invalid ad account status.');

  const data = {
    name: clean_(payload.name), email, contact: clean_(payload.contact), business: clean_(payload.business),
    facebookPageUrl: clean_(payload.facebookPageUrl), websiteUrl: clean_(payload.websiteUrl),
    businessPortfolioStatus: clean_(payload.businessPortfolioStatus), adsAccountStatus: clean_(payload.adsAccountStatus),
    previousAds: clean_(payload.previousAds), objective: clean_(payload.objective), productService: clean_(payload.productService),
    offer: clean_(payload.offer), targetLocations: clean_(payload.targetLocations), targetAudience: clean_(payload.targetAudience),
    dailyBudget: clean_(payload.dailyBudget), runDuration: clean_(payload.runDuration), creativeStatus: clean_(payload.creativeStatus),
    referenceLink: clean_(payload.referenceLink), notes: clean_(payload.notes), package: clean_(payload.package)
  };

  validateLength_(data.name, 'Name', 2, 100);
  validateLength_(data.email, 'Email', 5, 150);
  validateLength_(data.contact, 'Contact', 3, 150);
  validateLength_(data.business, 'Business name', 2, 150);
  validateLength_(data.facebookPageUrl, 'Facebook Page URL', 0, 1000);
  validateLength_(data.websiteUrl, 'Website URL', 0, 1000);
  validateLength_(data.previousAds, 'Previous ads', 1, 150);
  validateLength_(data.objective, 'Objective', 1, 200);
  validateLength_(data.productService, 'Product or service', 3, 1500);
  validateLength_(data.offer, 'Offer', 2, 1000);
  validateLength_(data.targetLocations, 'Target locations', 2, 500);
  validateLength_(data.targetAudience, 'Target audience', 3, 1500);
  validateLength_(data.dailyBudget, 'Daily budget', 1, 100);
  validateLength_(data.runDuration, 'Run duration', 1, 100);
  validateLength_(data.creativeStatus, 'Creative status', 1, 200);
  validateLength_(data.referenceLink, 'Reference link', 0, 1000);
  validateLength_(data.notes, 'Additional details', 0, 2000);
  validateLength_(data.package, 'Package', 0, 250);
  validateOptionalUrl_(data.facebookPageUrl, 'Facebook Page URL');
  validateOptionalUrl_(data.websiteUrl, 'Website URL');

  return withPublicWriteLock_(() => {
    enforceSubmissionProtection_('metaAds', email, [data.name, data.email, data.business, data.objective, data.productService, data.offer]);

    const now = formatDate_(new Date());
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.META_ADS);
    appendObjectRow_(sheet, META_ADS_HEADERS, {
      id: Utilities.getUuid(), createdAt: now, ...data,
      source: clean_(payload.source || 'Ads Success PH Meta Ads Intake').slice(0, 150),
      consent: 'Yes', privacyConsent: 'Yes', status: 'New', updatedAt: now
    });

    markSubmissionProtection_('metaAds', email, [data.name, data.email, data.business, data.objective, data.productService, data.offer]);
    return { message: 'Detailed Meta Ads request saved.' };
  });
}


function submitReview_(payload) {
  if (clean_(payload.companyWebsite)) return { message: 'Review received.' };

  const required = ['name', 'email', 'service', 'rating', 'review', 'consent', 'privacyConsent'];
  required.forEach(key => {
    if (!String(payload[key] || '').trim()) throw new Error(`Missing required field: ${key}`);
  });

  const email = clean_(payload.email).toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Please enter a valid email address.');

  const rating = Number(payload.rating);
  if (!Number.isInteger(rating) || rating < 1 || rating > 5) throw new Error('Rating must be from 1 to 5.');

  const name = clean_(payload.name);
  const business = clean_(payload.business);
  const service = clean_(payload.service);
  const reviewText = clean_(payload.review);

  validateLength_(name, 'Name', 2, 80);
  validateLength_(email, 'Email', 5, 120);
  validateLength_(business, 'Business name', 0, 100);
  validateLength_(service, 'Service', 2, 100);
  validateLength_(reviewText, 'Review', 10, 1000);

  if (payload.consent !== 'yes') throw new Error('Review publishing consent is required.');
  if (payload.privacyConsent !== 'yes') throw new Error('Please agree to the Privacy Notice before submitting.');

  return withPublicWriteLock_(() => {
    enforceSubmissionProtection_('review', email, [name, email, service, rating, reviewText]);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.REVIEWS);
    const now = formatDate_(new Date());
    appendObjectRow_(sheet, REVIEW_HEADERS, {
      id: Utilities.getUuid(), createdAt: now, name, email, business, service,
      rating, review: reviewText, privacyConsent: 'Yes', status: 'Pending', updatedAt: now
    });

    markSubmissionProtection_('review', email, [name, email, service, rating, reviewText]);
    return { message: 'Review saved and waiting for approval.' };
  });
}

function adminLogin_(payload) {
  const storedHash = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD_HASH');
  if (!storedHash) throw new Error('Admin password is not configured yet.');
  if (hash_(String(payload.password || '')) !== storedHash) throw new Error('Incorrect admin password.');

  const token = Utilities.getUuid();
  CacheService.getScriptCache().put(`admin:${token}`, '1', 21600);
  return { token };
}

function uploadDemoImage_(payload) {
  return uploadPublicImage_(payload, 'demo');
}


function uploadPackageImage_(payload) {
  return uploadPublicImage_(payload, 'package');
}

function uploadPublicImage_(payload, prefix) {
  const mimeType = clean_(payload.mimeType);
  const fileName = clean_(payload.fileName);
  const base64 = clean_(payload.base64);

  if (!mimeType.startsWith('image/')) throw new Error('Only image files are allowed.');
  if (!base64) throw new Error('Image data is missing.');

  let bytes;
  try {
    bytes = Utilities.base64Decode(base64);
  } catch (error) {
    throw new Error('Unable to decode the uploaded image.');
  }

  if (bytes.length > MAX_UPLOAD_BYTES) throw new Error('Image must be 4 MB or smaller.');

  const safeName = sanitizeFileName_(fileName || `${prefix}-${Date.now()}.jpg`);
  const blob = Utilities.newBlob(bytes, mimeType, `${prefix}-${Date.now()}-${safeName}`);
  const folder = getOrCreateUploadFolder_();
  const file = folder.createFile(blob);

  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    try { file.setSecurityUpdateEnabled(false); } catch (ignored) {}
  } catch (error) {
    file.setTrashed(true);
    throw new Error('Google Drive could not make the image publicly viewable. Check your Drive sharing policy or use a Google account that allows link sharing.');
  }

  return {
    fileId: file.getId(),
    imageUrl: `https://drive.google.com/thumbnail?id=${file.getId()}&sz=w1600`
  };
}


function saveDemo_(payload) {
  const required = ['title', 'category', 'description', 'imageUrl', 'demoUrl'];
  required.forEach(key => {
    if (!String(payload[key] || '').trim()) throw new Error(`Missing required field: ${key}`);
  });

  const now = formatDate_(new Date());
  const item = {
    id: clean_(payload.id) || Utilities.getUuid(),
    title: clean_(payload.title),
    category: clean_(payload.category),
    description: clean_(payload.description),
    imageUrl: clean_(payload.imageUrl),
    demoUrl: clean_(payload.demoUrl),
    createdAt: now,
    updatedAt: now
  };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.DEMOS);
  upsertObjectRow_(sheet, DEMO_HEADERS, item, 'id', true);
  return { id: item.id };
}

function deleteDemo_(payload) {
  deleteObjectRow_(SHEETS.DEMOS, 'id', payload.id, 'Demo not found.');
  return { message: 'Demo deleted.' };
}

function savePackage_(payload) {
  const required = ['category', 'name', 'price', 'description', 'inclusions', 'active', 'featured'];
  required.forEach(key => {
    if (!String(payload[key] || '').trim()) throw new Error(`Missing required field: ${key}`);
  });

  const category = clean_(payload.category);
  if (!PACKAGE_CATEGORIES.includes(category)) throw new Error('Invalid package category.');

  const active = clean_(payload.active);
  const featured = clean_(payload.featured);
  if (!['Yes', 'No'].includes(active)) throw new Error('Invalid public visibility value.');
  if (!['Yes', 'No'].includes(featured)) throw new Error('Invalid featured value.');

  const sortOrder = Number(payload.sortOrder || 0);
  if (!Number.isFinite(sortOrder) || sortOrder < 0) throw new Error('Sort order must be zero or higher.');

  let sampleImageUrls = [];
  const rawSamples = clean_(payload.sampleImageUrls);
  if (rawSamples) {
    try {
      const parsed = JSON.parse(rawSamples);
      if (!Array.isArray(parsed)) throw new Error('not array');
      sampleImageUrls = parsed.map(clean_).filter(Boolean);
    } catch (error) {
      throw new Error('Sample image list is invalid.');
    }
  }
  if (sampleImageUrls.length > 8) throw new Error('A package can have a maximum of 8 sample images.');

  const now = formatDate_(new Date());
  const item = {
    id: clean_(payload.id) || Utilities.getUuid(),
    category,
    name: clean_(payload.name).slice(0, 100),
    price: clean_(payload.price).slice(0, 80),
    description: clean_(payload.description).slice(0, 500),
    inclusions: clean_(payload.inclusions).slice(0, 3000),
    badge: clean_(payload.badge).slice(0, 50),
    ctaLabel: clean_(payload.ctaLabel).slice(0, 50),
    coverImageUrl: clean_(payload.coverImageUrl).slice(0, 1500),
    sampleImageUrls: JSON.stringify(sampleImageUrls),
    active,
    featured,
    sortOrder: Math.floor(sortOrder),
    createdAt: now,
    updatedAt: now
  };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PACKAGES);
  upsertObjectRow_(sheet, PACKAGE_HEADERS, item, 'id', true);
  return { id: item.id };
}

function deletePackage_(payload) {
  deleteObjectRow_(SHEETS.PACKAGES, 'id', payload.id, 'Package not found.');
  return { message: 'Package deleted.' };
}

function updateLeadStatus_(payload) {
  const allowed = ['New', 'Contacted', 'Quoted', 'Closed', 'Not Proceeding'];
  if (!allowed.includes(payload.status)) throw new Error('Invalid lead status.');
  updateObjectField_(SHEETS.LEADS, 'id', payload.id, 'status', payload.status, 'Lead not found.');
  return { message: 'Lead status updated.' };
}


function updateMetaAdsStatus_(payload) {
  const allowed = ['New', 'Reviewing', 'Waiting for Client', 'Waiting for Access', 'Ready to Launch', 'Active', 'Completed', 'Not Proceeding'];
  if (!allowed.includes(payload.status)) throw new Error('Invalid Meta Ads request status.');

  updateObjectField_(SHEETS.META_ADS, 'id', payload.id, 'status', payload.status, 'Meta Ads request not found.');
  updateObjectField_(SHEETS.META_ADS, 'id', payload.id, 'updatedAt', formatDate_(new Date()), 'Meta Ads request not found.');
  return { message: 'Meta Ads request status updated.' };
}


function updateReviewStatus_(payload) {
  const allowed = ['Pending', 'Approved', 'Rejected'];
  if (!allowed.includes(payload.status)) throw new Error('Invalid review status.');

  updateObjectField_(SHEETS.REVIEWS, 'id', payload.id, 'status', payload.status, 'Review not found.');
  updateObjectField_(SHEETS.REVIEWS, 'id', payload.id, 'updatedAt', formatDate_(new Date()), 'Review not found.');
  return { message: 'Review status updated.' };
}

function deleteReview_(payload) {
  deleteObjectRow_(SHEETS.REVIEWS, 'id', payload.id, 'Review not found.');
  return { message: 'Review deleted.' };
}

function saveSettings_(payload) {
  const allowed = ['brandName', 'logoUrl', 'heroTitle', 'heroSubtitle', 'messengerUrl'];
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.SETTINGS);
  const data = sheet.getDataRange().getValues();

  allowed.forEach(key => {
    const value = clean_(payload[key]);
    const rowIndex = data.findIndex((row, index) => index > 0 && String(row[0]) === key);

    if (rowIndex === -1) {
      sheet.appendRow([key, value]);
    } else {
      sheet.getRange(rowIndex + 1, 2).setValue(value);
    }
  });

  return { message: 'Settings saved.' };
}

function getSettingsObject_() {
  const rows = getRowsAsObjects_(SHEETS.SETTINGS);
  return rows.reduce((acc, item) => {
    acc[item.key] = item.value;
    return acc;
  }, {});
}

function getOrCreateUploadFolder_() {
  const properties = PropertiesService.getScriptProperties();
  const folderId = properties.getProperty('UPLOAD_FOLDER_ID');

  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (ignored) {
      properties.deleteProperty('UPLOAD_FOLDER_ID');
    }
  }

  const existing = DriveApp.getFoldersByName(UPLOAD_FOLDER_NAME);
  const folder = existing.hasNext() ? existing.next() : DriveApp.createFolder(UPLOAD_FOLDER_NAME);
  properties.setProperty('UPLOAD_FOLDER_ID', folder.getId());
  return folder;
}

function getRowsAsObjects_(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Missing sheet: ${sheetName}. Run setupApp() first.`);

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return [];
  const headers = values[0];

  return values.slice(1)
    .filter(row => row.some(cell => String(cell).trim() !== ''))
    .map(row => headers.reduce((obj, header, index) => {
      obj[header] = row[index] || '';
      return obj;
    }, {}));
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const lastColumn = Math.max(sheet.getLastColumn(), 1);
    const currentHeaders = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0]
      .map(value => String(value).trim());

    const duplicateHeaders = currentHeaders.filter((header, index) => header && currentHeaders.indexOf(header) !== index);
    if (duplicateHeaders.length) throw new Error(`Sheet "${name}" contains duplicate header names.`);

    headers.forEach(header => {
      if (!currentHeaders.includes(header)) {
        sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
        currentHeaders.push(header);
      }
    });
  }

  sheet.setFrozenRows(1);
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  headerRange.setFontWeight('bold');
  headerRange.setNote('SYSTEM DATABASE HEADER — Do not rename, delete, or reorder columns unless the Apps Script code is updated.');
  protectHeaderRow_(sheet, headerRange);
}

function appendObjectRow_(sheet, headers, object) {
  const actualHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const row = actualHeaders.map(header => Object.prototype.hasOwnProperty.call(object, header) ? object[header] : '');
  sheet.appendRow(row);
}

function upsertObjectRow_(sheet, headers, object, keyHeader, preserveCreatedAt) {
  const values = sheet.getDataRange().getValues();
  const actualHeaders = values[0].map(value => String(value));
  const keyIndex = actualHeaders.indexOf(keyHeader);
  if (keyIndex === -1) throw new Error(`Missing key header: ${keyHeader}`);

  const rowIndex = values.findIndex((row, index) =>
    index > 0 && String(row[keyIndex]) === String(object[keyHeader])
  );

  if (rowIndex === -1) {
    appendObjectRow_(sheet, headers, object);
    return;
  }

  if (preserveCreatedAt && actualHeaders.includes('createdAt')) {
    const createdIndex = actualHeaders.indexOf('createdAt');
    object.createdAt = values[rowIndex][createdIndex] || object.createdAt;
  }

  const newRow = actualHeaders.map((header, index) =>
    Object.prototype.hasOwnProperty.call(object, header) ? object[header] : values[rowIndex][index]
  );
  sheet.getRange(rowIndex + 1, 1, 1, newRow.length).setValues([newRow]);
}

function deleteObjectRow_(sheetName, keyHeader, keyValue, notFoundMessage) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(value => String(value));
  const keyIndex = headers.indexOf(keyHeader);
  const rowIndex = values.findIndex((row, index) => index > 0 && String(row[keyIndex]) === String(keyValue));
  if (rowIndex === -1) throw new Error(notFoundMessage);
  sheet.deleteRow(rowIndex + 1);
}

function updateObjectField_(sheetName, keyHeader, keyValue, fieldHeader, fieldValue, notFoundMessage) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(value => String(value));
  const keyIndex = headers.indexOf(keyHeader);
  const fieldIndex = headers.indexOf(fieldHeader);
  if (keyIndex === -1 || fieldIndex === -1) throw new Error(`Missing sheet header in ${sheetName}.`);

  const rowIndex = values.findIndex((row, index) => index > 0 && String(row[keyIndex]) === String(keyValue));
  if (rowIndex === -1) throw new Error(notFoundMessage);

  sheet.getRange(rowIndex + 1, fieldIndex + 1).setValue(fieldValue);
}

function withPublicWriteLock_(callback) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(PUBLIC_WRITE_LOCK_TIMEOUT_MS)) {
    throw new Error('The system is receiving many submissions right now. Please wait a moment and submit again.');
  }
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function enforceSubmissionProtection_(kind, identity, fingerprintParts) {
  const cache = CacheService.getScriptCache();
  const identityKey = submissionIdentityKey_(kind, identity);
  const duplicateKey = submissionDuplicateKey_(kind, fingerprintParts);
  if (cache.get(identityKey)) {
    const seconds = PUBLIC_COOLDOWN_SECONDS[kind] || 30;
    throw new Error(`Please wait about ${seconds} seconds before sending another ${submissionLabel_(kind)}.`);
  }
  if (cache.get(duplicateKey)) {
    throw new Error('This looks like a duplicate submission. Your first request may already have been received.');
  }
}

function markSubmissionProtection_(kind, identity, fingerprintParts) {
  const cache = CacheService.getScriptCache();
  cache.put(submissionIdentityKey_(kind, identity), '1', PUBLIC_COOLDOWN_SECONDS[kind] || 30);
  cache.put(submissionDuplicateKey_(kind, fingerprintParts), '1', PUBLIC_DUPLICATE_TTL_SECONDS[kind] || 120);
}

function submissionIdentityKey_(kind, identity) {
  return `rate:${kind}:${hash_(clean_(identity).toLowerCase())}`;
}

function submissionDuplicateKey_(kind, parts) {
  const normalized = parts.map(value => clean_(value).toLowerCase()).join('|');
  return `dedupe:${kind}:${hash_(normalized)}`;
}

function submissionLabel_(kind) {
  if (kind === 'metaAds') return 'Meta Ads request';
  if (kind === 'review') return 'review';
  return 'service request';
}

function validateLength_(value, label, minLength, maxLength) {
  const length = clean_(value).length;
  if (length < minLength) throw new Error(`${label} must be at least ${minLength} characters.`);
  if (length > maxLength) throw new Error(`${label} must be ${maxLength} characters or less.`);
}

function validateOptionalUrl_(value, label) {
  const url = clean_(value);
  if (!url) return;
  if (!/^https?:\/\//i.test(url)) throw new Error(`${label} must start with http:// or https://.`);
}

function protectHeaderRow_(sheet, headerRange) {
  const description = `Ads Success PH system header: ${sheet.getName()}`;
  const existing = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .find(protection => protection.getDescription() === description);
  const protection = existing || headerRange.protect();
  protection.setDescription(description);
  protection.setRange(headerRange);
  protection.setWarningOnly(true);
}

function requireAdminToken_(token) {
  if (!token || CacheService.getScriptCache().get(`admin:${token}`) !== '1') {
    throw new Error('Admin session expired. Please log in again.');
  }
}

function parsePayload_(e) {
  const raw = e.parameter && e.parameter.payload;
  if (!raw) return {};
  try {
    return JSON.parse(raw);
  } catch (error) {
    throw new Error('Invalid request payload.');
  }
}

function clean_(value) {
  return String(value == null ? '' : value).trim();
}

function sanitizeFileName_(value) {
  const cleaned = clean_(value).replace(/[^a-zA-Z0-9._-]+/g, '-').replace(/-+/g, '-');
  return cleaned.slice(0, 120) || 'demo-image';
}

function hash_(value) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8);
  return bytes.map(byte => {
    const v = byte < 0 ? byte + 256 : byte;
    return ('0' + v.toString(16)).slice(-2);
  }).join('');
}

function formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function json_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
