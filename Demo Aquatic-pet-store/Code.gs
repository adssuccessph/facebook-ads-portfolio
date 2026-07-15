const ADMIN_PASSCODE = "Riri@09991249908";
const ROOT_FOLDER = "Kawaii Aqua Pets Uploads";
const LOGO_FOLDER = "Store Logos";
const PRODUCT_FOLDER = "Product Images";
const ORDER_FOLDER = "Payment Proofs";
const REVIEW_FOLDER = "Review Images";
const PAYMENT_QR_FOLDER = "Payment QR Codes";
const DEFAULT_LALAMOVE_FORM_URL = "https://delivery.lalamove.com/forms/PH4c4ef013d6d54893b979fa6c04c447ca";
const LBC_TRACKING_URL = "https://www.lbcexpress.com/track/";

function doGet() {
  return out({
    ok: true,
    message: "Kawaii Aqua Pets API online",
    time: new Date()
  });
}

function doPost(e) {
  try {
    const r = JSON.parse((e.postData && e.postData.contents) || "{}");
    const action = r.action;

    const adminActions = [
      "adminGetAll",
      "saveSettings",
      "saveCategory",
      "deleteCategory",
      "saveProduct",
      "deleteProduct",
      "updateOrderStatus",
      "saveTrackingInfo",
      "deleteReview"
    ];

    if (adminActions.includes(action) && r.passcode !== ADMIN_PASSCODE) {
      throw Error("Invalid admin passcode");
    }

    const handlers = {
      getStore,
      getPaymentQr,
      trackOrder,
      createOrder,
      createReview,
      adminGetAll,
      saveSettings,
      saveCategory,
      deleteCategory,
      saveProduct,
      deleteProduct,
      updateOrderStatus,
      saveTrackingInfo,
      deleteReview
    };

    if (!handlers[action]) throw Error("Unknown action");
    return out(Object.assign({ok: true}, handlers[action](r)));
  } catch (err) {
    return out({
      ok: false,
      error: String(err.message || err)
    });
  }
}

function setupStore() {
  mk("Settings", ["key", "value"]);
  mk("Categories", ["id", "name", "createdAt"]);
  mk("Products", [
    "id",
    "name",
    "categoryId",
    "price",
    "stock",
    "description",
    "imageUrl",
    "imageFileId",
    "active",
    "createdAt",
    "updatedAt"
  ]);
  mk("Orders", [
    "id",
    "createdAt",
    "customerName",
    "mobile",
    "email",
    "address",
    "notes",
    "itemsJson",
    "itemsSummary",
    "total",
    "proofUrl",
    "proofFileId",
    "status",
    "stockDeducted",
    "deliveryMethod",
    "deliveryJson",
    "deliverySummary",
    "lalamoveFormCompleted",
    "trackingNumber",
    "trackingUrl",
    "statusHistoryJson",
    "shippedAt",
    "completedAt",
    "cancelledAt",
    "shipmentEmailSentAt",
    "shipmentEmailStatus",
    "paymentMethod",
    "paymentSummary",
    "paidAt",
    "receiptEmailSentAt",
    "receiptEmailStatus",
    "clientRequestId"
  ]);
  mk("Reviews", [
    "id",
    "createdAt",
    "customerName",
    "productId",
    "productName",
    "rating",
    "reviewText",
    "imageUrl",
    "imageFileId",
    "clientRequestId"
  ]);

  ensureOrderColumns();

  const settingsSheet = sh("Settings");
  if (settingsSheet.getLastRow() === 1) {
    settingsSheet.getRange(2, 1, 19, 2).setValues([
      ["siteName", "Kawaii Aqua Pets"],
      ["tagline", "Quality Betta, Guppy & Aqua Pets • Nationwide Shipping"],
      ["gcashName", "Joebert Greganda"],
      ["gcashNumber", ""],
      ["gcashQrUrl", ""],
      ["gcashQrFileId", ""],
      ["unionBankName", "JOEBERT O GREGANDA"],
      ["unionBankAccountHint", "**** **** 6628"],
      ["unionBankQrUrl", ""],
      ["unionBankQrFileId", ""],
      ["paymentNote", "Using one phone only? Download the QR first, then select the saved QR image inside your payment app."],
      ["logoUrl", ""],
      ["lalamoveFormUrl", DEFAULT_LALAMOVE_FORM_URL],
      ["senderName", ""],
      ["senderMobile", ""],
      ["senderAddress", ""],
      ["facebookPageUrl", ""],
      ["tiktokUrl", ""],
      ["youtubeUrl", ""]
    ]);
  }

  seedDemoStoreData();

  rootFolder();

  // Running setupStore once also authorizes the email service used for paid-order receipts.
  MailApp.getRemainingDailyQuota();
}


function seedDemoStoreData() {
  const categoryDefs = [
    {key: "betta", id: "DEMO_CAT_BETTA", name: "Betta", aliases: ["betta", "betta fish"]},
    {key: "guppies", id: "DEMO_CAT_GUPPIES", name: "Guppies", aliases: ["guppy", "guppies"]},
    {key: "snails", id: "DEMO_CAT_SNAILS", name: "Snails", aliases: ["snail", "snails"]},
    {key: "shrimps", id: "DEMO_CAT_SHRIMPS", name: "Shrimps", aliases: ["shrimp", "shrimps"]}
  ];

  const categoriesSheet = sh("Categories");
  let currentCategories = rows("Categories");
  const categoryIds = {};

  categoryDefs.forEach(def => {
    const found = currentCategories.find(category =>
      def.aliases.includes(demoKey_(category.name))
    );

    if (found) {
      categoryIds[def.key] = found.id;
      if (String(found.name || "").trim() !== def.name) {
        upd(categoriesSheet, found.id, {name: def.name});
      }
    } else {
      categoriesSheet.appendRow([def.id, def.name, new Date()]);
      categoryIds[def.key] = def.id;
      currentCategories = rows("Categories");
    }
  });

  const demoProducts = [
    ["DEMO_PRD_BETTA_001", "Galaxy Koi Betta Male", "betta", 350, 8, "Bright galaxy-koi pattern with bold red, blue, and marble tones. Individually kept and carefully packed.", "#e54b64", "#3157d7", 1],
    ["DEMO_PRD_BETTA_002", "Dumbo Halfmoon Betta", "betta", 420, 5, "Large dumbo fins with a wide halfmoon tail. A premium centerpiece betta for planted tanks.", "#f4c6ff", "#7b4bd9", 2],
    ["DEMO_PRD_BETTA_003", "Avatar Blue Betta", "betta", 480, 4, "Deep electric-blue body with dark contrast and metallic shine. Limited demo stock.", "#1f4fd6", "#38c7e8", 3],
    ["DEMO_PRD_BETTA_004", "Nemo Candy Betta", "betta", 390, 7, "Colorful nemo candy mix with orange, red, white, and speckled markings. Each fish has a unique pattern.", "#ff7a3d", "#ffe06b", 4],
    ["DEMO_PRD_BETTA_005", "Samurai Black Betta", "betta", 450, 3, "Dark samurai-style betta with metallic scale contrast. Best for hobbyists who like dramatic colors.", "#20242a", "#9aa4b2", 5],
    ["DEMO_PRD_BETTA_006", "Female Betta Sorority Pack", "betta", 650, 6, "Demo pack of assorted female bettas for experienced keepers planning a properly sized sorority setup.", "#ef6aa8", "#f4b5d2", 6],

    ["DEMO_PRD_GUPPY_001", "Blue Moscow Guppy Pair", "guppies", 220, 12, "Healthy male and female pair with rich blue Moscow coloration. Great starter breeding pair.", "#245edb", "#4ec5e8", 7],
    ["DEMO_PRD_GUPPY_002", "Full Red Guppy Trio", "guppies", 280, 10, "One male and two females with strong full-red color. Active, hardy, and suitable for breeding projects.", "#d93131", "#ff7f61", 8],
    ["DEMO_PRD_GUPPY_003", "Purple Dragon Guppy Pair", "guppies", 260, 8, "Purple dragon pattern with detailed body scales and flowing tails. Carefully selected demo pair.", "#7e45c9", "#d886ef", 9],
    ["DEMO_PRD_GUPPY_004", "Koi Tuxedo Guppy Trio", "guppies", 300, 6, "Koi head color with tuxedo body contrast. Trio setup is ideal for hobbyist breeding tanks.", "#ff6b4a", "#343a55", 10],
    ["DEMO_PRD_GUPPY_005", "Albino Full Platinum Guppy Pair", "guppies", 320, 5, "Light platinum body with albino eyes and clean metallic finish. Premium-looking pair for display tanks.", "#f6f1de", "#d8d4c8", 11],

    ["DEMO_PRD_SNAIL_001", "Nerite Snail 3 pcs", "snails", 120, 20, "Three assorted nerite snails. Popular algae grazers for freshwater aquariums.", "#6b5135", "#c59b62", 12],
    ["DEMO_PRD_SNAIL_002", "Mystery Snail Gold", "snails", 80, 18, "Bright golden mystery snail with a cute round shell. Peaceful addition to community tanks.", "#f0b930", "#ffe790", 13],
    ["DEMO_PRD_SNAIL_003", "Pink Ramshorn Snail 5 pcs", "snails", 100, 15, "Five pink ramshorn snails with soft rose shell tones. Great for planted aquarium clean-up crews.", "#ee8fb3", "#ffd1df", 14],
    ["DEMO_PRD_SNAIL_004", "Assassin Snail", "snails", 90, 12, "Striped assassin snail often kept by hobbyists managing pest snail populations.", "#d7a33d", "#4e3825", 15],

    ["DEMO_PRD_SHRIMP_001", "Red Cherry Shrimp 10 pcs", "shrimps", 280, 20, "Ten active red cherry shrimps for planted nano tanks. Great beginner shrimp colony starter.", "#e63232", "#ff8173", 16],
    ["DEMO_PRD_SHRIMP_002", "Blue Dream Shrimp 10 pcs", "shrimps", 380, 12, "Ten blue dream shrimps with deep blue body color. Best viewed on light substrate and green plants.", "#3159c9", "#4cc8e6", 17],
    ["DEMO_PRD_SHRIMP_003", "Yellow Goldenback Shrimp 10 pcs", "shrimps", 350, 10, "Ten yellow goldenback shrimps with bright dorsal color. Eye-catching in planted aquascapes.", "#f2c72e", "#fff08a", 18],
    ["DEMO_PRD_SHRIMP_004", "Amano Shrimp 5 pcs", "shrimps", 300, 14, "Five amano shrimps known among aquarists as active algae grazers and tank clean-up helpers.", "#879b8a", "#d0ddd3", 19],
    ["DEMO_PRD_SHRIMP_005", "Crystal Red Shrimp 5 pcs", "shrimps", 450, 6, "Five crystal red shrimps with classic red-and-white banding. Recommended for stable mature tanks.", "#d92f3d", "#f8f5ef", 20]
  ];

  const productsSheet = sh("Products");
  let existingProducts = rows("Products");
  const existingProductNames = new Set(
    existingProducts.map(product => demoKey_(product.name))
  );

  demoProducts.forEach(product => {
    const [id, name, categoryKey, price, stock, description, colorA, colorB, variant] = product;
    if (existingProductNames.has(demoKey_(name))) return;

    appendByHeaders(productsSheet, {
      id,
      name,
      categoryId: categoryIds[categoryKey],
      price,
      stock,
      description,
      imageUrl: demoProductImageDataUrl_(name, categoryDefs.find(def => def.key === categoryKey).name, colorA, colorB, variant),
      imageFileId: "",
      active: true,
      createdAt: new Date(),
      updatedAt: new Date()
    });

    existingProductNames.add(demoKey_(name));
  });

  existingProducts = rows("Products");
  const productByName = Object.fromEntries(
    existingProducts.map(product => [demoKey_(product.name), product])
  );

  const demoReviews = [
    ["REV_DEMO_001", 3, "Ana M.", "Galaxy Koi Betta Male", 5, "DEMO REVIEW — Ang ganda ng color ng betta at maayos ang sample packaging flow. Perfect pang-test ng review layout."],
    ["REV_DEMO_002", 5, "Carlo D.", "Blue Moscow Guppy Pair", 5, "DEMO REVIEW — Active tingnan yung pair at malinaw ang product details. Ang dali rin i-filter sa Guppies category."],
    ["REV_DEMO_003", 7, "Mika R.", "Red Cherry Shrimp 10 pcs", 5, "DEMO REVIEW — Cute ng shrimp listing! Malinis tingnan sa mobile at obvious agad ang price at stock."],
    ["REV_DEMO_004", 9, "Jessa P.", "Mystery Snail Gold", 4, "DEMO REVIEW — Maganda yung category layout at mabilis makita ang snails. Nice demo product card."],
    ["REV_DEMO_005", 11, "Ruel S.", "Nemo Candy Betta", 5, "DEMO REVIEW — Gusto ko yung Buy Now at Add to Cart buttons. Hindi nakakalito kahit maraming products."],
    ["REV_DEMO_006", 13, "Bea L.", "Blue Dream Shrimp 10 pcs", 5, "DEMO REVIEW — Premium tingnan yung Blue Dream listing. Malinis at bagay sa aquatic store theme."],
    ["REV_DEMO_007", 15, "Kevin T.", "Full Red Guppy Trio", 5, "DEMO REVIEW — Ang dali mag-browse. Helpful yung available stock na nakalagay sa bawat card."],
    ["REV_DEMO_008", 17, "Mae G.", "Nerite Snail 3 pcs", 4, "DEMO REVIEW — Simple pero complete yung product info. Gusto ko rin na may customer reviews section."],
    ["REV_DEMO_009", 19, "Jonas B.", "Avatar Blue Betta", 5, "DEMO REVIEW — Solid yung color theme at responsive sa phone. Mukhang actual online aqua shop na."],
    ["REV_DEMO_010", 21, "Trisha C.", "Yellow Goldenback Shrimp 10 pcs", 5, "DEMO REVIEW — Madaling makita ang Shrimps category at malinaw ang presyo. Good customer-view demo."],
    ["REV_DEMO_011", 23, "Paolo N.", "Samurai Black Betta", 5, "DEMO REVIEW — Clean product grid at hindi crowded. Malakas yung dating ng product name at stock info."],
    ["REV_DEMO_012", 25, "Elle F.", "Pink Ramshorn Snail 5 pcs", 4, "DEMO REVIEW — Cute ng aquatic theme. Okay yung spacing at readable kahit maraming review cards."],
    ["REV_DEMO_013", 27, "Mark A.", "Koi Tuxedo Guppy Trio", 5, "DEMO REVIEW — Ganda ng category chips. Isang click lang, filtered agad ang guppies."],
    ["REV_DEMO_014", 29, "Rica V.", "Amano Shrimp 5 pcs", 5, "DEMO REVIEW — Mas mukhang established shop dahil may products at feedback agad sa customer view."],
    ["REV_DEMO_015", 31, "Dennis Q.", "Dumbo Halfmoon Betta", 5, "DEMO REVIEW — Smooth yung product browsing at maayos yung card design sa desktop."],
    ["REV_DEMO_016", 34, "Faith O.", "Crystal Red Shrimp 5 pcs", 5, "DEMO REVIEW — Nice demo catalog. Malinaw din na hiwalay ang products, categories, at reviews."],
    ["REV_DEMO_017", 38, "Leo H.", "Albino Full Platinum Guppy Pair", 4, "DEMO REVIEW — Maganda pang presentation sa client dahil hindi blank ang storefront."],
    ["REV_DEMO_018", 42, "Kim S.", "Assassin Snail", 5, "DEMO REVIEW — Kumpleto tingnan ang customer page. Helpful din yung Track Order button sa header."]
  ];

  const reviewsSheet = sh("Reviews");
  const existingReviewIds = new Set(
    rows("Reviews").map(review => String(review.id || ""))
  );

  demoReviews.forEach(review => {
    const [id, daysAgo, customerName, productName, rating, reviewText] = review;
    if (existingReviewIds.has(id)) return;

    const product = productByName[demoKey_(productName)] || {};
    reviewsSheet.appendRow([
      id,
      new Date(Date.now() - daysAgo * 24 * 60 * 60 * 1000),
      customerName,
      product.id || "",
      product.name || productName,
      rating,
      reviewText,
      "",
      ""
    ]);

    existingReviewIds.add(id);
  });

  return {
    categories: categoryDefs.length,
    demoProducts: demoProducts.length,
    demoReviews: demoReviews.length
  };
}

function removeDemoStoreData() {
  removeRowsWithIdPrefix_("Reviews", "REV_DEMO_");
  removeRowsWithIdPrefix_("Products", "DEMO_PRD_");
  return {removed: true};
}

function removeRowsWithIdPrefix_(sheetName, prefix) {
  const sheet = sh(sheetName);
  if (sheet.getLastRow() < 2) return;

  for (let rowNumber = sheet.getLastRow(); rowNumber >= 2; rowNumber--) {
    const id = String(sheet.getRange(rowNumber, 1).getValue() || "");
    if (id.indexOf(prefix) === 0) sheet.deleteRow(rowNumber);
  }
}

function demoKey_(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function demoProductImageDataUrl_(name, categoryName, colorA, colorB, variant) {
  const type = demoKey_(categoryName);
  const spotShift = Number(variant || 1) * 9;
  let subject = "";

  if (type === "snails") {
    subject =
      '<ellipse cx="425" cy="365" rx="180" ry="54" fill="' + colorB + '" opacity=".95"/>' +
      '<circle cx="405" cy="315" r="122" fill="' + colorA + '"/>' +
      '<circle cx="405" cy="315" r="78" fill="none" stroke="#ffffff" stroke-width="18" opacity=".55"/>' +
      '<circle cx="405" cy="315" r="35" fill="none" stroke="#ffffff" stroke-width="12" opacity=".45"/>' +
      '<circle cx="575" cy="338" r="18" fill="#ffffff"/>' +
      '<circle cx="582" cy="338" r="8" fill="#162b32"/>' +
      '<path d="M565 310 Q575 260 596 244 M586 312 Q606 270 630 263" stroke="#415d58" stroke-width="8" stroke-linecap="round"/>';
  } else if (type === "shrimps") {
    subject =
      '<path d="M275 355 Q365 215 535 285 Q625 320 600 430 Q570 535 405 475" fill="none" stroke="' + colorA + '" stroke-width="74" stroke-linecap="round"/>' +
      '<path d="M320 330 Q390 260 500 295" fill="none" stroke="' + colorB + '" stroke-width="18" stroke-linecap="round" opacity=".9"/>' +
      '<circle cx="563" cy="318" r="15" fill="#ffffff"/>' +
      '<circle cx="568" cy="319" r="7" fill="#132b31"/>' +
      '<path d="M610 338 Q680 285 725 300 M610 356 Q690 350 730 390" fill="none" stroke="' + colorA + '" stroke-width="12" stroke-linecap="round"/>' +
      '<path d="M345 410 l-60 80 M405 445 l-35 92 M470 455 l5 86" stroke="' + colorA + '" stroke-width="12" stroke-linecap="round"/>';
  } else {
    subject =
      '<polygon points="250,350 115,245 125,455" fill="' + colorB + '" opacity=".95"/>' +
      '<ellipse cx="445" cy="350" rx="220" ry="128" fill="' + colorA + '"/>' +
      '<path d="M400 245 Q475 150 560 235 Q505 260 455 305" fill="' + colorB + '" opacity=".85"/>' +
      '<path d="M395 455 Q475 540 565 452 Q500 430 445 395" fill="' + colorB + '" opacity=".82"/>' +
      '<circle cx="585" cy="322" r="22" fill="#ffffff"/>' +
      '<circle cx="592" cy="325" r="10" fill="#10231f"/>' +
      '<circle cx="' + (330 + spotShift) + '" cy="325" r="28" fill="' + colorB + '" opacity=".65"/>' +
      '<circle cx="' + (420 + spotShift) + '" cy="385" r="21" fill="#ffffff" opacity=".45"/>' +
      '<circle cx="' + (505 - spotShift) + '" cy="305" r="18" fill="' + colorB + '" opacity=".6"/>';
  }

  const svg =
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 900 700">' +
      '<defs>' +
        '<linearGradient id="bg" x1="0" y1="0" x2="1" y2="1">' +
          '<stop offset="0" stop-color="#dff7f4"/>' +
          '<stop offset="1" stop-color="#f7fbff"/>' +
        '</linearGradient>' +
      '</defs>' +
      '<rect width="900" height="700" rx="48" fill="url(#bg)"/>' +
      '<circle cx="105" cy="115" r="38" fill="#ffffff" opacity=".6"/>' +
      '<circle cx="760" cy="145" r="24" fill="#ffffff" opacity=".7"/>' +
      '<circle cx="805" cy="240" r="14" fill="#ffffff" opacity=".8"/>' +
      '<path d="M40 520 Q140 455 220 525 T410 515 T610 525 T860 505 V700 H40Z" fill="#bfe9df" opacity=".72"/>' +
      '<path d="M75 540 Q110 445 135 540 M155 540 Q205 425 220 540 M710 540 Q740 435 765 540 M775 540 Q820 460 842 540" stroke="#58a995" stroke-width="18" stroke-linecap="round" fill="none" opacity=".72"/>' +
      subject +
      '<rect x="54" y="535" width="792" height="118" rx="28" fill="#ffffff" opacity=".94"/>' +
      '<text x="84" y="582" font-family="Arial,sans-serif" font-size="24" font-weight="700" fill="#087f67">' + htmlEscape(categoryName).toUpperCase() + '</text>' +
      '<text x="84" y="626" font-family="Arial,sans-serif" font-size="35" font-weight="800" fill="#10231f">' + htmlEscape(name) + '</text>' +
    '</svg>';

  return "data:image/svg+xml;charset=UTF-8," + encodeURIComponent(svg);
}

function ensureOrderColumns() {
  [
    "stockDeducted",
    "deliveryMethod",
    "deliveryJson",
    "deliverySummary",
    "lalamoveFormCompleted",
    "trackingNumber",
    "trackingUrl",
    "statusHistoryJson",
    "shippedAt",
    "completedAt",
    "cancelledAt",
    "shipmentEmailSentAt",
    "shipmentEmailStatus",
    "paymentMethod",
    "paymentSummary",
    "email",
    "paidAt",
    "receiptEmailSentAt",
    "receiptEmailStatus",
    "clientRequestId"
  ].forEach(header => ensureCol("Orders", header));

  ensureCol("Reviews", "clientRequestId");
}

function readSettings() {
  const settings = {};
  rows("Settings").forEach(row => settings[row.key] = row.value);

  if (!settings.siteName || settings.siteName === "Pojhes Pet Store") {
    settings.siteName = "Kawaii Aqua Pets";
  }
  if (
    !settings.tagline ||
    settings.tagline === "Quality Betta Fish • Nationwide Shipping"
  ) {
    settings.tagline = "Quality Betta, Guppy & Aqua Pets • Nationwide Shipping";
  }
  if (!settings.lalamoveFormUrl) settings.lalamoveFormUrl = DEFAULT_LALAMOVE_FORM_URL;
  if (
    !settings.gcashName ||
    settings.gcashName === "JO****T G." ||
    settings.gcashName === "Pojhes Pet Store"
  ) {
    settings.gcashName = "Joebert Greganda";
  }
  if (settings.gcashNumber == null || settings.gcashNumber === "09XX XXX XXXX") {
    settings.gcashNumber = "";
  }
  if (!settings.unionBankName) settings.unionBankName = "JOEBERT O GREGANDA";
  if (!settings.unionBankAccountHint) settings.unionBankAccountHint = "**** **** 6628";
  if (!settings.paymentNote) settings.paymentNote = "Using one phone only? Download the QR first, then select the saved QR image inside your payment app.";
  if (!settings.facebookPageUrl) settings.facebookPageUrl = "";
  if (!settings.tiktokUrl) settings.tiktokUrl = "";
  if (!settings.youtubeUrl) settings.youtubeUrl = "";

  return settings;
}

function getStore() {
  const settings = readSettings();
  settings.logoUrl = publicImageUrl(settings.logoFileId || settings.logoUrl || "");
  settings.gcashQrUrl = publicImageUrl(settings.gcashQrFileId || settings.gcashQrUrl || "");
  settings.unionBankQrUrl = publicImageUrl(settings.unionBankQrFileId || settings.unionBankQrUrl || "");

  const categories = rows("Categories");
  const categoryMap = Object.fromEntries(categories.map(category => [category.id, category.name]));

  const products = rows("Products")
    .map(product => Object.assign(product, {
      price: Number(product.price),
      stock: Number(product.stock),
      categoryName: categoryMap[product.categoryId] || "",
      active: String(product.active) !== "false",
      imageUrl: publicImageUrl(product.imageFileId || product.imageUrl)
    }))
    .filter(product => product.active);

  const reviews = rows("Reviews")
    .map(review => Object.assign(review, {
      rating: Number(review.rating),
      imageUrl: publicImageUrl(review.imageFileId || review.imageUrl),
      isDemo: String(review.id || "").indexOf("REV_DEMO_") === 0
    }))
    .reverse();

  return {
    settings,
    categories,
    products,
    reviews
  };
}

function adminGetAll() {
  ensureOrderColumns();

  const store = getStore();
  const categoryMap = Object.fromEntries(
    store.categories.map(category => [category.id, category.name])
  );

  const products = rows("Products").map(product => Object.assign(product, {
    categoryName: categoryMap[product.categoryId] || "",
    price: Number(product.price),
    stock: Number(product.stock),
    imageUrl: publicImageUrl(product.imageFileId || product.imageUrl)
  }));

  const orders = rows("Orders")
    .reverse()
    .map(order => Object.assign(order, {
      proofUrl: publicImageUrl(order.proofFileId || order.proofUrl)
    }));

  const reviews = rows("Reviews")
    .map(review => Object.assign(review, {
      rating: Number(review.rating),
      imageUrl: publicImageUrl(review.imageFileId || review.imageUrl),
      isDemo: String(review.id || "").indexOf("REV_DEMO_") === 0
    }))
    .reverse();

  return {
    settings: store.settings,
    categories: store.categories,
    products,
    orders,
    reviews
  };
}

function saveSettings(r) {
  const current = readSettings();
  const settings = r.settings || {};

  if (settings.logoData) {
    const oldId = current.logoFileId || extractDriveId(current.logoUrl);
    const image = saveImage(
      settings.logoData,
      settings.logoName || "logo.jpg",
      LOGO_FOLDER
    );

    current.logoUrl = image.url;
    current.logoFileId = image.id;
    trashFile(oldId, image.id);
  }

  if (settings.gcashQrData) {
    const oldId = current.gcashQrFileId || extractDriveId(current.gcashQrUrl);
    const image = saveImage(
      settings.gcashQrData,
      settings.gcashQrName || "gcash-qr.jpg",
      PAYMENT_QR_FOLDER
    );
    current.gcashQrUrl = image.url;
    current.gcashQrFileId = image.id;
    trashFile(oldId, image.id);
  }

  if (settings.unionBankQrData) {
    const oldId = current.unionBankQrFileId || extractDriveId(current.unionBankQrUrl);
    const image = saveImage(
      settings.unionBankQrData,
      settings.unionBankQrName || "unionbank-qr.jpg",
      PAYMENT_QR_FOLDER
    );
    current.unionBankQrUrl = image.url;
    current.unionBankQrFileId = image.id;
    trashFile(oldId, image.id);
  }

  [
    "siteName",
    "tagline",
    "gcashName",
    "gcashNumber",
    "unionBankName",
    "unionBankAccountHint",
    "paymentNote",
    "lalamoveFormUrl",
    "senderName",
    "senderMobile",
    "senderAddress",
    "facebookPageUrl",
    "tiktokUrl",
    "youtubeUrl"
  ].forEach(key => {
    current[key] = String(settings[key] || "");
  });

  if (!current.lalamoveFormUrl) {
    current.lalamoveFormUrl = DEFAULT_LALAMOVE_FORM_URL;
  }

  writeSettings(current);

  return {
    logoUrl: publicImageUrl(current.logoFileId || current.logoUrl),
    gcashQrUrl: publicImageUrl(current.gcashQrFileId || current.gcashQrUrl),
    unionBankQrUrl: publicImageUrl(current.unionBankQrFileId || current.unionBankQrUrl)
  };
}


function getPaymentQr(r) {
  const method = String(r.method || "").toLowerCase();
  const settings = readSettings();
  const isUnionBank = method === "unionbank";
  const fileId = isUnionBank
    ? (settings.unionBankQrFileId || extractDriveId(settings.unionBankQrUrl))
    : (settings.gcashQrFileId || extractDriveId(settings.gcashQrUrl));

  if (!fileId) return {dataUrl: "", filename: ""};

  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  const mime = blob.getContentType() || "image/jpeg";
  const bytes = blob.getBytes();

  if (!mime.startsWith("image/")) throw Error("Stored payment QR is not an image");
  if (bytes.length > 6 * 1024 * 1024) throw Error("Payment QR image is too large to download");

  const ext = mime === "image/png" ? "png" : mime === "image/webp" ? "webp" : "jpg";
  return {
    dataUrl: "data:" + mime + ";base64," + Utilities.base64Encode(bytes),
    filename: "Kawaii-Aqua-Pets-" + (isUnionBank ? "UnionBank" : "GCash") + "-QR." + ext
  };
}


function trackOrder(r) {
  ensureOrderColumns();

  const orderId = clean(r.orderId).toUpperCase();
  const contact = clean(r.contact);

  if (!orderId) throw Error("Order ID is required");
  if (!contact) throw Error("Email or mobile number is required");

  const order = rows("Orders").find(row =>
    String(row.id || "").toUpperCase() === orderId
  );

  if (!order || !contactMatches(order, contact)) {
    throw Error("Order not found. Check your Order ID and email or mobile number.");
  }

  const items = parseItems(order).map(item => ({
    name: item.name,
    qty: Math.max(1, Math.floor(Number(item.qty) || 1)),
    price: Number(item.price) || 0
  }));

  return {
    order: {
      id: order.id,
      createdAt: order.createdAt,
      items,
      itemsSummary: order.itemsSummary || "",
      total: Number(order.total) || 0,
      status: clean(order.status) || "Pending",
      paymentSummary: clean(order.paymentSummary) || clean(order.paymentMethod),
      deliveryMethod: clean(order.deliveryMethod).toLowerCase(),
      deliverySummary: clean(order.deliverySummary),
      trackingNumber: clean(order.trackingNumber),
      trackingUrl: publicTrackingUrl(order),
      statusHistory: getStatusHistory(order)
    }
  };
}

function contactMatches(order, contact) {
  const input = clean(contact).toLowerCase();
  const orderEmail = clean(order.email).toLowerCase();

  if (input.includes("@")) {
    return !!orderEmail && input === orderEmail;
  }

  const inputPhone = normalizePhone(input);
  const orderPhone = normalizePhone(order.mobile);

  return !!inputPhone && !!orderPhone && inputPhone === orderPhone;
}

function normalizePhone(value) {
  const digits = String(value || "").replace(/\D/g, "");
  return digits.length >= 10 ? digits.slice(-10) : digits;
}

function safeHttpUrl(value) {
  const url = clean(value);
  return /^https?:\/\/[^\s]+$/i.test(url) ? url : "";
}

function publicTrackingUrl(order) {
  const savedUrl = safeHttpUrl(order.trackingUrl);
  if (savedUrl) return savedUrl;

  if (
    clean(order.deliveryMethod).toLowerCase() === "lbc" &&
    clean(order.trackingNumber)
  ) {
    return LBC_TRACKING_URL;
  }

  return "";
}

function getStatusHistory(order) {
  let history = [];

  try {
    const parsed = JSON.parse(order.statusHistoryJson || "[]");
    if (Array.isArray(parsed)) history = parsed;
  } catch (err) {}

  history = history
    .filter(event => event && clean(event.status))
    .map(event => ({
      status: clean(event.status),
      at: clean(event.at)
    }));

  const addIfMissing = (status, at) => {
    if (!at || history.some(event => event.status === status)) return;
    const date = at instanceof Date ? at : new Date(at);
    history.push({
      status,
      at: Number.isNaN(date.getTime()) ? "" : date.toISOString()
    });
  };

  addIfMissing("Pending", order.createdAt);
  addIfMissing("Paid", order.paidAt);
  addIfMissing("Shipped", order.shippedAt);
  addIfMissing("Completed", order.completedAt);
  addIfMissing("Cancelled", order.cancelledAt);

  const currentStatus = clean(order.status) || "Pending";
  if (!history.some(event => event.status === currentStatus)) {
    history.push({status: currentStatus, at: ""});
  }

  return history.slice(-30);
}

function appendStatusHistory(order, status, at) {
  const history = getStatusHistory(order);
  const date = at instanceof Date ? at : new Date(at || new Date());
  const iso = Number.isNaN(date.getTime()) ? new Date().toISOString() : date.toISOString();

  if (history.length && history[history.length - 1].status === status) {
    if (!history[history.length - 1].at) history[history.length - 1].at = iso;
    return JSON.stringify(history.slice(-30));
  }

  history.push({status, at: iso});
  return JSON.stringify(history.slice(-30));
}

function saveTrackingInfo(r) {
  ensureOrderColumns();

  const order = rows("Orders").find(row => row.id == r.id);
  if (!order) throw Error("Order not found");

  const trackingNumber = clean(r.trackingNumber).slice(0, 120);
  const rawTrackingUrl = clean(r.trackingUrl);
  const trackingUrl = safeHttpUrl(rawTrackingUrl);

  if (rawTrackingUrl && !trackingUrl) {
    throw Error("Tracking link must start with http:// or https://");
  }

  upd(sh("Orders"), r.id, {
    trackingNumber,
    trackingUrl
  });

  const updatedOrder = Object.assign({}, order, {
    trackingNumber,
    trackingUrl
  });

  return {
    trackingNumber,
    trackingUrl: publicTrackingUrl(updatedOrder)
  };
}

function saveCategory(r) {
  const category = r.category || {};
  const name = String(category.name || "").trim();

  if (!name) throw Error("Category name is required");

  if (category.id) {
    upd(sh("Categories"), category.id, {name});
  } else {
    sh("Categories").appendRow([uid("cat"), name, new Date()]);
  }

  return {};
}

function deleteCategory(r) {
  if (rows("Products").some(product => product.categoryId == r.id)) {
    throw Error("Move or delete products first");
  }

  del(sh("Categories"), r.id);
  return {};
}

function saveProduct(r) {
  const product = r.product || {};

  if (!product.name) throw Error("Product name is required");
  if (!product.categoryId) throw Error("Category is required");

  const productsSheet = sh("Products");
  const now = new Date();
  let image = {};

  if (product.imageData) {
    image = saveImage(
      product.imageData,
      product.imageName || "product.jpg",
      PRODUCT_FOLDER
    );
  }

  if (product.id) {
    const oldProduct = rows("Products").find(row => row.id == product.id);
    if (!oldProduct) throw Error("Product not found");

    upd(productsSheet, product.id, {
      name: String(product.name).trim(),
      categoryId: product.categoryId,
      price: Number(product.price),
      stock: Math.max(0, Math.floor(Number(product.stock) || 0)),
      description: String(product.description || ""),
      imageUrl: image.url || oldProduct.imageUrl,
      imageFileId: image.id || oldProduct.imageFileId,
      active: !!product.active,
      updatedAt: now
    });

    if (image.id) {
      trashFile(
        oldProduct.imageFileId || extractDriveId(oldProduct.imageUrl),
        image.id
      );
    }
  } else {
    if (!image.id) throw Error("Product image is required");

    productsSheet.appendRow([
      uid("prd"),
      String(product.name).trim(),
      product.categoryId,
      Number(product.price),
      Math.max(0, Math.floor(Number(product.stock) || 0)),
      String(product.description || ""),
      image.url,
      image.id,
      !!product.active,
      now,
      now
    ]);
  }

  return {};
}

function deleteProduct(r) {
  const product = rows("Products").find(row => row.id == r.id);

  if (product) {
    trashFile(product.imageFileId || extractDriveId(product.imageUrl));
  }

  del(sh("Products"), r.id);
  return {};
}

function createOrder(r) {
  ensureOrderColumns();

  const order = r.order || {};
  const clientRequestId = clean(order.clientRequestId).slice(0, 160);

  if (!order.items || !order.items.length) throw Error("Cart is empty");
  if (!order.proofData) throw Error("Proof of payment is required");
  if (!clean(order.customerName)) throw Error("Customer name is required");
  if (!clean(order.mobile)) throw Error("Mobile number is required");
  if (!isValidEmail(order.email)) throw Error("A valid email address is required for the receipt");
  if (!clean(order.address)) throw Error("Complete address is required");
  if (!clientRequestId) throw Error("Checkout request ID is missing. Please refresh the store and try again.");

  const paymentMethod = String(order.paymentMethod || "").toLowerCase();
  if (!["gcash", "unionbank"].includes(paymentMethod)) {
    throw Error("Please select GCash or UnionBank as the payment method");
  }
  const paymentSummary = paymentMethod === "unionbank" ? "UnionBank" : "GCash";
  const delivery = buildDeliveryData(order);

  // Save proof before entering the sheet lock so multiple customers do not spend
  // unnecessary time waiting while Drive processes image uploads.
  const proof = saveImage(
    order.proofData,
    order.proofName || "proof.jpg",
    ORDER_FOLDER
  );

  const lock = LockService.getScriptLock();
  let lockAcquired = false;
  let committed = false;

  try {
    lockAcquired = lock.tryLock(60000);
    if (!lockAcquired) {
      throw Error("Store is busy processing other orders. Please try again in a few seconds.");
    }

    const existingOrder = rows("Orders").find(row =>
      clean(row.clientRequestId) === clientRequestId
    );

    if (existingOrder) {
      trashFile(proof.id);
      return orderResponse_(existingOrder, true);
    }

    const currentProducts = rows("Products");
    const validItems = [];
    let total = 0;

    order.items.forEach(item => {
      const product = currentProducts.find(row => row.id == item.id);
      if (!product) throw Error(item.name + " is no longer available");

      const quantity = Math.max(1, Math.floor(Number(item.qty) || 1));
      const stock = Math.max(0, Number(product.stock) || 0);

      if (quantity > stock) {
        throw Error(product.name + ": only " + stock + " item(s) available");
      }

      const price = Number(product.price) || 0;

      validItems.push({
        id: product.id,
        name: product.name,
        price,
        qty: quantity,
        imageUrl: publicImageUrl(product.imageFileId || product.imageUrl)
      });

      total += price * quantity;
    });

    const id = uid("ORD");
    const createdAt = new Date();
    const itemsSummary = validItems
      .map(item => item.name + " x" + item.qty)
      .join(", ");

    appendByHeaders(sh("Orders"), {
      id,
      createdAt,
      customerName: clean(order.customerName),
      mobile: clean(order.mobile),
      email: clean(order.email).toLowerCase(),
      address: clean(order.address),
      notes: clean(order.notes),
      itemsJson: JSON.stringify(validItems),
      itemsSummary,
      total,
      proofUrl: proof.url,
      proofFileId: proof.id,
      status: "Pending",
      stockDeducted: false,
      deliveryMethod: delivery.method,
      deliveryJson: JSON.stringify(delivery.details),
      deliverySummary: delivery.summary,
      lalamoveFormCompleted: delivery.lalamoveFormCompleted,
      trackingNumber: "",
      trackingUrl: "",
      statusHistoryJson: JSON.stringify([{status: "Pending", at: createdAt.toISOString()}]),
      shippedAt: "",
      completedAt: "",
      cancelledAt: "",
      shipmentEmailSentAt: "",
      shipmentEmailStatus: "Waiting for shipment",
      paymentMethod,
      paymentSummary,
      paidAt: "",
      receiptEmailSentAt: "",
      receiptEmailStatus: "Waiting for payment approval",
      clientRequestId
    });

    SpreadsheetApp.flush();
    committed = true;

    return {
      orderId: id,
      total,
      itemsSummary,
      deliverySummary: delivery.summary,
      paymentSummary,
      status: "Pending",
      duplicate: false
    };
  } catch (err) {
    if (!committed) trashFile(proof.id);
    throw err;
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

function orderResponse_(order, duplicate) {
  return {
    orderId: clean(order.id),
    total: Number(order.total) || 0,
    itemsSummary: clean(order.itemsSummary),
    deliverySummary: clean(order.deliverySummary),
    paymentSummary: clean(order.paymentSummary) || (clean(order.paymentMethod) === "unionbank" ? "UnionBank" : "GCash"),
    status: clean(order.status) || "Pending",
    duplicate: !!duplicate
  };
}

function buildDeliveryData(order) {
  const method = String(order.deliveryMethod || "").trim().toLowerCase();

  if (method === "lalamove") {
    if (!toBool(order.lalamoveFormCompleted)) {
      throw Error("Please complete the Lalamove delivery form and confirm it before submitting.");
    }

    const settings = readSettings();

    return {
      method: "lalamove",
      summary: "Lalamove • Customer confirmed delivery form completed",
      lalamoveFormCompleted: true,
      details: {
        method: "Lalamove",
        formUrl: settings.lalamoveFormUrl || DEFAULT_LALAMOVE_FORM_URL,
        formCompleted: true
      }
    };
  }

  if (method === "lbc") {
    const lbc = order.lbc || {};
    const serviceType = clean(lbc.serviceType) === "Branch Pickup"
      ? "Branch Pickup"
      : "Door-to-Door";

    requireText(lbc.receiverName, "LBC receiver full name is required");
    requireText(lbc.mobile, "LBC receiver mobile number is required");

    const details = {
      method: "LBC",
      receiverName: clean(lbc.receiverName),
      mobile: clean(lbc.mobile),
      email: clean(lbc.email),
      serviceType,
      province: clean(lbc.province),
      cityMunicipality: clean(lbc.cityMunicipality),
      barangay: clean(lbc.barangay),
      postalCode: clean(lbc.postalCode),
      houseUnit: clean(lbc.houseUnit),
      streetSubdivision: clean(lbc.streetSubdivision),
      landmark: clean(lbc.landmark),
      branchProvince: clean(lbc.branchProvince),
      branchCity: clean(lbc.branchCity),
      branchName: clean(lbc.branchName),
      validIdName: clean(lbc.validIdName),
      instructions: clean(lbc.instructions)
    };

    let summary = "";

    if (serviceType === "Branch Pickup") {
      requireText(details.branchProvince, "LBC branch province is required");
      requireText(details.branchCity, "LBC branch city or municipality is required");
      requireText(details.branchName, "Preferred LBC branch is required");
      requireText(details.validIdName, "Receiver name on valid ID is required");

      summary = "LBC • Branch Pickup • " + details.branchName;
    } else {
      requireText(details.province, "LBC province is required");
      requireText(details.cityMunicipality, "LBC city or municipality is required");
      requireText(details.barangay, "LBC barangay is required");
      requireText(details.postalCode, "LBC postal or ZIP code is required");
      requireText(details.houseUnit, "LBC house, unit, lot, or block number is required");
      requireText(details.streetSubdivision, "LBC street or subdivision is required");

      summary = "LBC • Door-to-Door • " +
        details.cityMunicipality + ", " + details.province;
    }

    return {
      method: "lbc",
      summary,
      lalamoveFormCompleted: false,
      details
    };
  }

  throw Error("Please select Lalamove or LBC as the delivery method");
}

function createReview(r) {
  ensureOrderColumns();

  const review = r.review || {};
  const clientRequestId = clean(review.clientRequestId).slice(0, 160);
  const name = clean(review.customerName);
  const text = clean(review.reviewText);
  const rating = Math.max(
    1,
    Math.min(5, Math.floor(Number(review.rating) || 5))
  );

  if (!clientRequestId) throw Error("Review request ID is missing. Please refresh and try again.");
  if (!name) throw Error("Customer name is required");
  if (!text) throw Error("Review text is required");
  if (name.length > 60) throw Error("Customer name is too long");
  if (text.length > 800) throw Error("Review is too long");

  let productName = "";
  if (review.productId) {
    const product = rows("Products").find(row => row.id == review.productId);
    if (product) productName = product.name;
  }

  const image = review.imageData
    ? saveImage(
        review.imageData,
        review.imageName || "review.jpg",
        REVIEW_FOLDER
      )
    : {};

  const lock = LockService.getScriptLock();
  let lockAcquired = false;
  let committed = false;

  try {
    lockAcquired = lock.tryLock(60000);
    if (!lockAcquired) {
      throw Error("Store is busy processing other submissions. Please try again in a few seconds.");
    }

    const existingReview = rows("Reviews").find(row =>
      clean(row.clientRequestId) === clientRequestId
    );

    if (existingReview) {
      if (image.id) trashFile(image.id);
      return {reviewId: existingReview.id, duplicate: true};
    }

    const id = uid("REV");
    appendByHeaders(sh("Reviews"), {
      id,
      createdAt: new Date(),
      customerName: name,
      productId: review.productId || "",
      productName,
      rating,
      reviewText: text,
      imageUrl: image.url || "",
      imageFileId: image.id || "",
      clientRequestId
    });

    SpreadsheetApp.flush();
    committed = true;
    return {reviewId: id, duplicate: false};
  } catch (err) {
    if (!committed && image.id) trashFile(image.id);
    throw err;
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

function deleteReview(r) {
  const review = rows("Reviews").find(row => row.id == r.id);

  if (review) {
    trashFile(review.imageFileId || extractDriveId(review.imageUrl));
  }

  del(sh("Reviews"), r.id);
  return {};
}

function updateOrderStatus(r) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const order = rows("Orders").find(row => row.id == r.id);
    if (!order) throw Error("Order not found");

    const validStatuses = [
      "Pending",
      "Paid",
      "Preparing",
      "Ready",
      "Shipped",
      "Completed",
      "Cancelled"
    ];
    const status = clean(r.status);
    const currentStatus = clean(order.status) || "Pending";
    const stockDeducted = toBool(order.stockDeducted);

    if (!validStatuses.includes(status)) throw Error("Invalid order status");
    if (status === currentStatus) {
      return {inventoryAction: "none", emailNotification: null, shippingEmailNotification: null};
    }

    if (currentStatus === "Cancelled") {
      throw Error("A cancelled order cannot be reopened from the status menu.");
    }

    const rank = {
      Pending: 0,
      Paid: 1,
      Preparing: 2,
      Ready: 3,
      Shipped: 4,
      Completed: 5
    };

    if (
      status !== "Cancelled" &&
      currentStatus !== "Cancelled" &&
      Object.prototype.hasOwnProperty.call(rank, status) &&
      Object.prototype.hasOwnProperty.call(rank, currentStatus) &&
      rank[status] < rank[currentStatus]
    ) {
      throw Error("Order status cannot move backward. Use Cancelled when the order must be stopped.");
    }

    const now = new Date();

    if (status === "Paid" && !stockDeducted) {
      deductStock(order);

      upd(sh("Orders"), r.id, {
        status: "Paid",
        stockDeducted: true,
        paidAt: now,
        statusHistoryJson: appendStatusHistory(order, "Paid", now),
        receiptEmailStatus: "Preparing payment confirmation email"
      });

      const paidOrder = Object.assign({}, order, {
        status: "Paid",
        stockDeducted: true,
        paidAt: now,
        statusHistoryJson: appendStatusHistory(order, "Paid", now)
      });

      let emailNotification = {
        sent: false,
        status: "Skipped",
        message: "No receipt email was sent."
      };

      try {
        emailNotification = sendPaidReceiptEmail(paidOrder);

        upd(sh("Orders"), r.id, {
          receiptEmailSentAt: emailNotification.sent ? new Date() : "",
          receiptEmailStatus: emailNotification.message
        });
      } catch (emailErr) {
        emailNotification = {
          sent: false,
          status: "Failed",
          message: "Payment approved, but receipt email failed: " + String(emailErr.message || emailErr)
        };

        upd(sh("Orders"), r.id, {
          receiptEmailStatus: emailNotification.message
        });
      }

      return {
        inventoryAction: "deducted",
        emailNotification,
        shippingEmailNotification: null
      };
    }

    if (["Preparing", "Ready", "Shipped", "Completed"].includes(status) && !stockDeducted) {
      throw Error("Approve the payment first. Set the order to Paid.");
    }

    if (status === "Cancelled") {
      if (stockDeducted) restoreStock(order);

      upd(sh("Orders"), r.id, {
        status: "Cancelled",
        stockDeducted: false,
        cancelledAt: now,
        statusHistoryJson: appendStatusHistory(order, "Cancelled", now)
      });

      return {
        inventoryAction: stockDeducted ? "restored" : "none",
        emailNotification: null,
        shippingEmailNotification: null
      };
    }

    if (status === "Pending" && stockDeducted) {
      throw Error("A paid order cannot return to Pending.");
    }

    const patch = {
      status,
      statusHistoryJson: appendStatusHistory(order, status, now)
    };

    if (status === "Shipped") patch.shippedAt = now;
    if (status === "Completed") patch.completedAt = now;

    upd(sh("Orders"), r.id, patch);

    let shippingEmailNotification = null;

    if (status === "Shipped") {
      const shippedOrder = Object.assign({}, order, patch, {status: "Shipped"});

      try {
        shippingEmailNotification = sendShipmentEmail(shippedOrder);
        upd(sh("Orders"), r.id, {
          shipmentEmailSentAt: shippingEmailNotification.sent ? new Date() : "",
          shipmentEmailStatus: shippingEmailNotification.message
        });
      } catch (emailErr) {
        shippingEmailNotification = {
          sent: false,
          status: "Failed",
          message: "Order marked Shipped, but shipping email failed: " + String(emailErr.message || emailErr)
        };
        upd(sh("Orders"), r.id, {
          shipmentEmailStatus: shippingEmailNotification.message
        });
      }
    }

    return {
      inventoryAction: "none",
      emailNotification: null,
      shippingEmailNotification
    };
  } finally {
    lock.releaseLock();
  }
}

function deductStock(order) {
  const items = parseItems(order);
  const products = rows("Products");

  items.forEach(item => {
    const product = products.find(row => row.id == item.id);
    const quantity = Math.max(1, Math.floor(Number(item.qty) || 1));

    if (!product) {
      throw Error(item.name + ": product record no longer exists");
    }

    const stock = Math.max(0, Number(product.stock) || 0);

    if (stock < quantity) {
      throw Error(
        product.name +
        " cannot be approved. Only " +
        stock +
        " stock left but order needs " +
        quantity
      );
    }
  });

  items.forEach(item => {
    const product = rows("Products").find(row => row.id == item.id);
    const quantity = Math.max(1, Math.floor(Number(item.qty) || 1));

    upd(sh("Products"), item.id, {
      stock: Math.max(0, Number(product.stock) - quantity),
      updatedAt: new Date()
    });
  });
}

function restoreStock(order) {
  parseItems(order).forEach(item => {
    const product = rows("Products").find(row => row.id == item.id);
    if (!product) return;

    const quantity = Math.max(1, Math.floor(Number(item.qty) || 1));

    upd(sh("Products"), item.id, {
      stock: Math.max(0, Number(product.stock) || 0) + quantity,
      updatedAt: new Date()
    });
  });
}

function parseItems(order) {
  try {
    const items = JSON.parse(order.itemsJson || "[]");

    if (!Array.isArray(items) || !items.length) {
      throw Error();
    }

    return items;
  } catch (err) {
    throw Error("Order item data is invalid");
  }
}

function sendPaidReceiptEmail(order) {
  const recipient = clean(order.email).toLowerCase();

  if (!isValidEmail(recipient)) {
    return {
      sent: false,
      status: "Skipped",
      message: "Payment approved. No valid customer email is available for the receipt."
    };
  }

  if (MailApp.getRemainingDailyQuota() < 1) {
    throw Error("Daily Apps Script email recipient quota has been reached");
  }

  const settings = readSettings();
  const storeName = clean(settings.siteName) || "Kawaii Aqua Pets";
  const paymentLabel = clean(order.paymentSummary) ||
    (String(order.paymentMethod || "").toLowerCase() === "unionbank" ? "UnionBank" : "GCash");
  const deliveryLabel = clean(order.deliverySummary) || clean(order.deliveryMethod) || "To be confirmed";
  const items = parseItems(order);
  const total = Number(order.total) || 0;
  const paidDate = order.paidAt instanceof Date ? order.paidAt : new Date();
  const timezone = Session.getScriptTimeZone() || "Asia/Manila";
  const paidDateText = Utilities.formatDate(paidDate, timezone, "MMM d, yyyy h:mm a");

  const itemRowsHtml = items.map(item => {
    const qty = Math.max(1, Math.floor(Number(item.qty) || 1));
    const unitPrice = Number(item.price) || 0;
    const lineTotal = qty * unitPrice;

    return (
      "<tr>" +
        "<td style=\"padding:10px 8px;border-bottom:1px solid #e6efec;\">" + htmlEscape(item.name) + "</td>" +
        "<td style=\"padding:10px 8px;border-bottom:1px solid #e6efec;text-align:center;\">" + qty + "</td>" +
        "<td style=\"padding:10px 8px;border-bottom:1px solid #e6efec;text-align:right;\">" + formatPeso(unitPrice) + "</td>" +
        "<td style=\"padding:10px 8px;border-bottom:1px solid #e6efec;text-align:right;font-weight:700;\">" + formatPeso(lineTotal) + "</td>" +
      "</tr>"
    );
  }).join("");

  const plainItems = items.map(item => {
    const qty = Math.max(1, Math.floor(Number(item.qty) || 1));
    const unitPrice = Number(item.price) || 0;
    return "- " + item.name + " x" + qty + " = " + formatPeso(qty * unitPrice);
  }).join("\n");

  const subject = "Payment Confirmed & Order Received - " + order.id;
  const plainBody = [
    "Hi " + clean(order.customerName) + ",",
    "",
    "We received your order and confirmed your payment.",
    "",
    storeName + " PAYMENT RECEIPT",
    "Order ID: " + order.id,
    "Payment approved: " + paidDateText,
    "Payment method: " + paymentLabel,
    "Delivery: " + deliveryLabel,
    "",
    "ITEMS",
    plainItems,
    "",
    "PRODUCT TOTAL / AMOUNT PAID: " + formatPeso(total),
    "SHIPPING FEE: NOT INCLUDED",
    "",
    "The amount above covers the product order only. The shipping fee will be confirmed separately based on the selected delivery method.",
    "",
    "Thank you for your order!",
    storeName
  ].join("\n");

  const htmlBody =
    "<div style=\"margin:0;padding:24px;background:#f5f8f7;font-family:Arial,sans-serif;color:#10231f;\">" +
      "<div style=\"max-width:640px;margin:auto;background:#ffffff;border:1px solid #dce8e4;border-radius:18px;overflow:hidden;\">" +
        "<div style=\"padding:24px;background:#e9f8f5;text-align:center;\">" +
          "<div style=\"font-size:34px;\">🐠</div>" +
          "<h1 style=\"margin:8px 0 4px;color:#075f50;font-size:26px;\">" + htmlEscape(storeName) + "</h1>" +
          "<p style=\"margin:0;color:#52645e;\">Payment Confirmed &amp; Order Received</p>" +
        "</div>" +
        "<div style=\"padding:26px;\">" +
          "<p style=\"font-size:16px;line-height:1.6;margin-top:0;\">Hi <b>" + htmlEscape(order.customerName) + "</b>,</p>" +
          "<p style=\"font-size:16px;line-height:1.6;\">We received your order and confirmed your payment. Here is your payment receipt.</p>" +

          "<div style=\"margin:20px 0;padding:16px;background:#f5faf8;border:1px solid #dce8e4;border-radius:12px;\">" +
            "<table style=\"width:100%;border-collapse:collapse;font-size:14px;\">" +
              receiptMetaRow("Order ID", order.id) +
              receiptMetaRow("Payment Approved", paidDateText) +
              receiptMetaRow("Payment Method", paymentLabel) +
              receiptMetaRow("Delivery", deliveryLabel) +
              receiptMetaRow("Status", "PAID / PAYMENT APPROVED") +
            "</table>" +
          "</div>" +

          "<h2 style=\"font-size:18px;margin:24px 0 10px;\">Receipt</h2>" +
          "<table style=\"width:100%;border-collapse:collapse;font-size:14px;\">" +
            "<thead><tr style=\"background:#eef7f4;\">" +
              "<th style=\"padding:10px 8px;text-align:left;\">Item</th>" +
              "<th style=\"padding:10px 8px;text-align:center;\">Qty</th>" +
              "<th style=\"padding:10px 8px;text-align:right;\">Price</th>" +
              "<th style=\"padding:10px 8px;text-align:right;\">Total</th>" +
            "</tr></thead>" +
            "<tbody>" + itemRowsHtml + "</tbody>" +
          "</table>" +

          "<div style=\"margin-top:18px;padding:16px;border-radius:12px;background:#f5faf8;\">" +
            "<div style=\"display:flex;justify-content:space-between;gap:16px;font-size:16px;margin-bottom:8px;\">" +
              "<span>Product Total / Amount Paid</span><b>" + formatPeso(total) + "</b>" +
            "</div>" +
            "<div style=\"display:flex;justify-content:space-between;gap:16px;color:#b42318;font-weight:700;\">" +
              "<span>Shipping Fee</span><span>NOT INCLUDED</span>" +
            "</div>" +
          "</div>" +

          "<div style=\"margin-top:18px;padding:14px 16px;background:#fff7e8;border:1px solid #f0d39b;border-radius:10px;color:#6b4b14;line-height:1.55;\">" +
            "<b>Shipping fee is not included in this receipt.</b><br>" +
            "The amount above covers the product order only. The shipping fee will be confirmed separately based on your selected delivery method." +
          "</div>" +

          "<p style=\"margin:24px 0 0;line-height:1.6;\">Thank you for your order! 🐟<br><b>" + htmlEscape(storeName) + "</b></p>" +
        "</div>" +
      "</div>" +
    "</div>";

  MailApp.sendEmail({
    to: recipient,
    subject,
    body: plainBody,
    htmlBody,
    name: storeName
  });

  return {
    sent: true,
    status: "Sent",
    message: "Payment approved. Receipt email sent to " + recipient + "."
  };
}

function sendShipmentEmail(order) {
  const recipient = clean(order.email).toLowerCase();

  if (!isValidEmail(recipient)) {
    return {
      sent: false,
      status: "Skipped",
      message: "Order marked Shipped. No valid customer email is available for the shipping update."
    };
  }

  if (MailApp.getRemainingDailyQuota() < 1) {
    throw Error("Daily Apps Script email recipient quota has been reached");
  }

  const settings = readSettings();
  const storeName = clean(settings.siteName) || "Kawaii Aqua Pets";
  const deliveryLabel = clean(order.deliverySummary) || clean(order.deliveryMethod) || "Courier delivery";
  const trackingNumber = clean(order.trackingNumber);
  const trackingUrl = publicTrackingUrl(order);
  const subject = "Your order has been shipped - " + order.id;

  const trackingPlain = trackingNumber
    ? "Tracking Number: " + trackingNumber
    : "Tracking Number: Not provided yet";

  const plainBody = [
    "Hi " + clean(order.customerName) + ",",
    "",
    "Your " + storeName + " order has been marked as shipped.",
    "",
    "Order ID: " + order.id,
    "Courier / Delivery: " + deliveryLabel,
    trackingPlain,
    trackingUrl ? "Courier Tracking Page: " + trackingUrl : "",
    "",
    "You can also open the Kawaii Aqua Pets website and use Track Order with your Order ID plus your email or mobile number.",
    "",
    "Thank you!",
    storeName
  ].filter(Boolean).join("\n");

  const trackingButton = trackingUrl
    ? '<p style="margin:20px 0;"><a href="' + htmlEscape(trackingUrl) + '" style="display:inline-block;padding:12px 18px;background:#087f67;color:#fff;text-decoration:none;border-radius:10px;font-weight:700;">Open Courier Tracking</a></p>'
    : "";

  const htmlBody =
    '<div style="margin:0;padding:24px;background:#f5f8f7;font-family:Arial,sans-serif;color:#10231f;">' +
      '<div style="max-width:620px;margin:auto;background:#fff;border:1px solid #dce8e4;border-radius:18px;overflow:hidden;">' +
        '<div style="padding:24px;background:#e9f8f5;text-align:center;">' +
          '<div style="font-size:34px;">📦🐠</div>' +
          '<h1 style="margin:8px 0 4px;color:#075f50;font-size:25px;">Your order has been shipped!</h1>' +
          '<p style="margin:0;color:#52645e;">' + htmlEscape(storeName) + '</p>' +
        '</div>' +
        '<div style="padding:26px;">' +
          '<p style="font-size:16px;line-height:1.6;margin-top:0;">Hi <b>' + htmlEscape(order.customerName) + '</b>,</p>' +
          '<p style="font-size:16px;line-height:1.6;">Your order has been marked as shipped. Here are the delivery details currently saved by the seller.</p>' +
          '<div style="margin:20px 0;padding:16px;background:#f5faf8;border:1px solid #dce8e4;border-radius:12px;">' +
            '<table style="width:100%;border-collapse:collapse;font-size:14px;">' +
              receiptMetaRow("Order ID", order.id) +
              receiptMetaRow("Delivery", deliveryLabel) +
              receiptMetaRow("Tracking Number", trackingNumber || "Not provided yet") +
              receiptMetaRow("Status", "SHIPPED") +
            '</table>' +
          '</div>' +
          trackingButton +
          '<p style="line-height:1.6;color:#52645e;">You can also return to the Kawaii Aqua Pets website and use <b>Track Order</b> with your Order ID plus your email or mobile number.</p>' +
          '<p style="margin:24px 0 0;line-height:1.6;">Thank you! 🐟<br><b>' + htmlEscape(storeName) + '</b></p>' +
        '</div>' +
      '</div>' +
    '</div>';

  MailApp.sendEmail({
    to: recipient,
    subject,
    body: plainBody,
    htmlBody,
    name: storeName
  });

  return {
    sent: true,
    status: "Sent",
    message: "Shipping update email sent to " + recipient + "."
  };
}

function receiptMetaRow(label, value) {
  return (
    "<tr>" +
      "<td style=\"padding:6px 0;color:#65756f;\">" + htmlEscape(label) + "</td>" +
      "<td style=\"padding:6px 0;text-align:right;font-weight:700;\">" + htmlEscape(value) + "</td>" +
    "</tr>"
  );
}

function formatPeso(value) {
  const amount = Number(value) || 0;
  return "₱" + amount.toLocaleString("en-PH", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 2
  });
}

function htmlEscape(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function isValidEmail(value) {
  const email = clean(value);
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function clean(value) {
  return String(value || "").trim();
}

function requireText(value, message) {
  if (!clean(value)) throw Error(message);
}

function toBool(value) {
  return value === true || String(value).toLowerCase() === "true";
}

function saveImage(data, name, subfolder) {
  if (!data || !data.includes(",")) throw Error("Invalid image data");

  const parts = data.split(",");
  const mime = (parts[0].match(/data:(.*?);/) || [])[1] || "image/jpeg";

  if (!mime.startsWith("image/")) {
    throw Error("Only image files are allowed");
  }

  const bytes = Utilities.base64Decode(parts[1]);

  if (bytes.length > 6 * 1024 * 1024) {
    throw Error("Optimized image is too large. Please use a smaller image.");
  }

  const safeName = String(name || "image.jpg")
    .replace(/[^a-zA-Z0-9._-]/g, "_");

  const file = childFolder(subfolder).createFile(
    Utilities.newBlob(
      bytes,
      mime,
      Date.now() + "_" + safeName
    )
  );

  file.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );

  return {
    id: file.getId(),
    url: publicImageUrl(file.getId())
  };
}

function publicImageUrl(value) {
  const id = extractDriveId(value);

  return id
    ? "https://drive.google.com/thumbnail?id=" + id + "&sz=w1600"
    : String(value || "");
}

function extractDriveId(value) {
  const text = String(value || "").trim();

  if (/^[-\w]{20,}$/.test(text)) return text;

  const match = text.match(/(?:[?&]id=|\/d\/)([-\w]{20,})/);
  return match ? match[1] : "";
}

function trashFile(id, exceptId) {
  if (!id || id === exceptId) return;

  try {
    DriveApp.getFileById(id).setTrashed(true);
  } catch (err) {}
}

function rootFolder() {
  const iterator = DriveApp.getFoldersByName(ROOT_FOLDER);

  return iterator.hasNext()
    ? iterator.next()
    : DriveApp.createFolder(ROOT_FOLDER);
}

function childFolder(name) {
  const root = rootFolder();
  const iterator = root.getFoldersByName(name);

  return iterator.hasNext()
    ? iterator.next()
    : root.createFolder(name);
}

function writeSettings(object) {
  const settingsSheet = sh("Settings");

  if (settingsSheet.getLastRow() > 1) {
    settingsSheet
      .getRange(2, 1, settingsSheet.getLastRow() - 1, 2)
      .clearContent();
  }

  const values = Object.entries(object);

  if (values.length) {
    settingsSheet.getRange(2, 1, values.length, 2).setValues(values);
  }
}

function appendByHeaders(sheet, object) {
  const lastColumn = sheet.getLastColumn();
  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0];

  sheet.appendRow(headers.map(header =>
    Object.prototype.hasOwnProperty.call(object, header)
      ? object[header]
      : ""
  ));
}

function rows(name) {
  const data = sh(name).getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];

  return data
    .slice(1)
    .filter(row => row.some(value => value !== ""))
    .map(row => Object.fromEntries(
      headers.map((header, index) => [header, row[index]])
    ));
}

function sh(name) {
  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName(name);

  if (!sheet) throw Error("Run setupStore first");
  return sheet;
}

function mk(name, headers) {
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
}

function ensureCol(sheetName, header) {
  const sheet = sh(sheetName);
  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  if (!headers.includes(header)) {
    sheet
      .getRange(1, sheet.getLastColumn() + 1)
      .setValue(header);
  }
}

function uid(prefix) {
  return prefix + "_" + Utilities.getUuid().slice(0, 8) + "_" + Date.now();
}

function upd(sheet, id, patch) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rowIndex = data.findIndex((row, index) =>
    index && row[0] == id
  );

  if (rowIndex < 1) throw Error("Not found");

  Object.entries(patch).forEach(([key, value]) => {
    const columnIndex = headers.indexOf(key);

    if (columnIndex >= 0) {
      sheet
        .getRange(rowIndex + 1, columnIndex + 1)
        .setValue(value);
    }
  });
}

function del(sheet, id) {
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex((row, index) =>
    index && row[0] == id
  );

  if (rowIndex > 0) {
    sheet.deleteRow(rowIndex + 1);
  }
}

function out(object) {
  return ContentService
    .createTextOutput(JSON.stringify(object))
    .setMimeType(ContentService.MimeType.JSON);
}
