/**
 * PEMS – প্রোডাক্ট ক্যাটালগ (Google Sheets) + Google সার্চ
 * এই ফাইলটি Apps Script-এ নতুন ফাইল (Catalog.gs) হিসেবে যোগ করুন।
 * ক্যাটালগ থাকবে আপনার "PEMS Master" Sheet-এর নতুন "Catalog" ট্যাবে।
 */
const CAT_SHEET = 'Catalog';
const CAT_HEAD = ['key', 'barcode', 'name', 'image', 'category', 'updated'];

function norm_(s) { return String(s || '').toLowerCase().replace(/\s+/g, ' ').trim(); }

function catalogSheet_() {
  return ensureSheet_(SpreadsheetApp.getActive(), CAT_SHEET, CAT_HEAD);
}

function readCatalog_() {
  const sh = catalogSheet_(), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, CAT_HEAD.length).getValues().map(function (v) {
    return { key: String(v[0]), barcode: String(v[1] || ''), name: String(v[2] || ''), image: String(v[3] || ''), category: String(v[4] || '') };
  });
}

function pub_(r) { return { name: r.name, image: r.image, category: r.category, barcode: r.barcode }; }

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

function saveCatalog_(item) {
  item = item || {};
  const name = String(item.name || '').trim().slice(0, 120);
  if (!name) return false;
  const barcode = String(item.barcode || '').replace(/\D/g, '').slice(0, 40);
  const image = String(item.image || '').slice(0, 40000);
  const category = String(item.category || '').trim().slice(0, 40);
  const key = barcode || 'n:' + norm_(name);

  const lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    const sh = catalogSheet_(), n = sh.getLastRow();
    const data = n > 1 ? sh.getRange(2, 1, n - 1, CAT_HEAD.length).getValues() : [];
    let idx = -1, i;
    for (i = 0; i < data.length; i++) if (String(data[i][0]) === key) { idx = i; break; }
    if (idx < 0 && barcode) { // আগে শুধু নাম দিয়ে সেভ ছিল — এবার বারকোডসহ আপডেট
      const nk = 'n:' + norm_(name);
      for (i = 0; i < data.length; i++) if (String(data[i][0]) === nk) { idx = i; break; }
    }
    const now = new Date().toISOString();
    if (idx >= 0) {
      const old = data[idx];
      sh.getRange(idx + 2, 1, 1, CAT_HEAD.length).setValues([[key, barcode || String(old[1] || ''), name, image || String(old[3] || ''), category || String(old[4] || ''), now]]);
    } else {
      sh.appendRow([key, barcode, name, image, category, now]);
    }
    return true;
  } finally { lock.releaseLock(); }
}

/* Google Images (SerpAPI) — Script properties-এ SERPAPI_KEY থাকলে চালু হয় */
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
  cache.put(ck, JSON.stringify(items), 21600);
  return items;
}
