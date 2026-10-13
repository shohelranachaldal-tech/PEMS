/**
 * PEMS – প্রোডাক্ট ক্যাটালগ (Google Sheets) + Google সার্চ
 * ------------------------------------------------------------
 * এই ফাইলটি আপনার Apps Script প্রজেক্টে নতুন ফাইল (Catalog.gs) হিসেবে যোগ করুন।
 *
 * ১) doPost-এ, টোকেন যাচাইয়ের ঠিক পরে (listProducts/saveProduct যেখানে হ্যান্ডেল হয়) এই ৩ লাইন বসান:
 *
 *      const extra = catalogApi_(body);   // body = JSON.parse(e.postData.contents)
 *      if (extra) return ContentService.createTextOutput(JSON.stringify(extra))
 *                          .setMimeType(ContentService.MimeType.JSON);
 *
 * ২) (ঐচ্ছিক) Google ছবি/তথ্য খোঁজা চালু করতে:
 *    Project Settings → Script properties → SERPAPI_KEY = আপনার serpapi.com API key
 *    (Google-এর নিজস্ব Custom Search API নতুন ব্যবহারকারীদের জন্য বন্ধ, তাই SerpAPI-র
 *     Google Images ব্যবহার করা হয়েছে। ফ্রি প্ল্যানে মাসে ১০০টি সার্চ।)
 *    কী না দিলেও অ্যাপ চলবে — তখন অনলাইন খোঁজায় শুধু Open Facts ব্যবহার হবে।
 *
 * ৩) নতুন ডিপ্লয় করুন (Deploy → Manage deployments → Edit → New version)
 *    এবং চাওয়া হলে অনুমতি (UrlFetch / Spreadsheet) দিন।
 *
 * ক্যাটালগ নিজে থেকে "PEMS Product Catalog" নামে একটি Google Sheet তৈরি করে নেবে।
 * ------------------------------------------------------------
 */
const CAT_SHEET = 'Catalog';
const CAT_HEAD = ['key', 'barcode', 'name', 'image', 'category', 'updated'];

function catalogApi_(body) {
  try {
    switch (body.action) {
      case 'searchCatalog': return { ok: true, items: searchCatalog_(body.q, body.barcode) };
      case 'saveCatalog':   return { ok: true, saved: saveCatalog_(body.item) };
      case 'googleSearch':  return { ok: true, items: googleSearch_(body.q) };
    }
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  }
  return null; // অন্য অ্যাকশন — আপনার আগের কোডে যাবে
}

function norm_(s) { return String(s || '').toLowerCase().replace(/\s+/g, ' ').trim(); }

function catalogSheet_() {
  const props = PropertiesService.getScriptProperties();
  let ss = null, created = false;
  const id = props.getProperty('CATALOG_SS_ID');
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) {} }
  if (!ss) {
    ss = SpreadsheetApp.create('PEMS Product Catalog'); created = true;
    props.setProperty('CATALOG_SS_ID', ss.getId());
  }
  let sh = ss.getSheetByName(CAT_SHEET);
  if (!sh) {
    sh = created ? ss.getSheets()[0].setName(CAT_SHEET) : ss.insertSheet(CAT_SHEET);
    sh.getRange(1, 1, 1, CAT_HEAD.length).setValues([CAT_HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:B').setNumberFormat('@'); // বারকোডের শুরুর ০ যেন হারিয়ে না যায়
  }
  return sh;
}

function readCatalog_() {
  const sh = catalogSheet_(), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, CAT_HEAD.length).getValues().map(function (v) {
    return { key: String(v[0]), barcode: String(v[1] || ''), name: String(v[2] || ''), image: String(v[3] || ''), category: String(v[4] || '') };
  });
}

function searchCatalog_(q, barcode) {
  const rows = readCatalog_();
  const bc = String(barcode || '').replace(/\D/g, '');
  if (bc) return rows.filter(function (r) { return r.barcode === bc; }).slice(0, 3).map(pub_);
  const nq = norm_(q);
  if (nq.length < 2) return [];
  const hit = rows.filter(function (r) { return norm_(r.name).indexOf(nq) !== -1; });
  hit.sort(function (a, b) { return (norm_(a.name).indexOf(nq) === 0 ? 0 : 1) - (norm_(b.name).indexOf(nq) === 0 ? 0 : 1); });
  return hit.slice(0, 8).map(pub_);
}
function pub_(r) { return { name: r.name, image: r.image, category: r.category, barcode: r.barcode }; }

function saveCatalog_(item) {
  item = item || {};
  const name = String(item.name || '').trim().slice(0, 120);
  if (!name) return false;
  const barcode = String(item.barcode || '').replace(/\D/g, '');
  const image = String(item.image || '').slice(0, 40000);
  const category = String(item.category || '').trim().slice(0, 40);
  const key = barcode || 'n:' + norm_(name);

  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = catalogSheet_(), n = sh.getLastRow();
    const data = n > 1 ? sh.getRange(2, 1, n - 1, CAT_HEAD.length).getValues() : [];
    let idx = -1, i;
    for (i = 0; i < data.length; i++) if (String(data[i][0]) === key) { idx = i; break; }
    if (idx < 0 && barcode) { // আগে শুধু নাম দিয়ে সেভ ছিল — এবার বারকোডসহ আপডেট
      const nk = 'n:' + norm_(name);
      for (i = 0; i < data.length; i++) if (String(data[i][0]) === nk) { idx = i; break; }
    }
    const now = new Date();
    if (idx >= 0) {
      const old = data[idx];
      sh.getRange(idx + 2, 1, 1, CAT_HEAD.length).setValues([[key, barcode || String(old[1] || ''), name, image || old[3], category || old[4], now]]);
    } else {
      sh.appendRow([key, barcode, name, image, category, now]);
    }
    return true;
  } finally { lock.releaseLock(); }
}

function googleSearch_(q) {
  q = String(q || '').trim();
  if (q.length < 2) return [];
  const apiKey = PropertiesService.getScriptProperties().getProperty('SERPAPI_KEY');
  if (!apiKey) return [];
  const cache = CacheService.getScriptCache();
  const ck = 'gs:' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, norm_(q)));
  const hit = cache.get(ck);
  if (hit) return JSON.parse(hit);
  const res = UrlFetchApp.fetch('https://serpapi.com/search.json?engine=google_images&gl=bd&q=' +
    encodeURIComponent(q) + '&api_key=' + encodeURIComponent(apiKey), { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return [];
  const j = JSON.parse(res.getContentText());
  const items = (j.images_results || []).slice(0, 6).map(function (r) {
    return { name: String(r.title || '').slice(0, 120), image: r.thumbnail || r.original || '' };
  }).filter(function (x) { return x.name; });
  cache.put(ck, JSON.stringify(items), 21600); // ৬ ঘণ্টা — কোটা বাঁচাতে
  return items;
}
