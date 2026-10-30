/************************************************************
 * DATASERVICE.GS — Unified FM & LM Data Loader (Optimized)
 ************************************************************/
const CACHE_TTL_SECONDS = 600;
const GLOBAL_SHEET_CACHE = {};
let CACHED_SS_INSTANCE = null;

function getSS() {
  if (!CACHED_SS_INSTANCE) {
    CACHED_SS_INSTANCE = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }
  return CACHED_SS_INSTANCE;
}

function readSheetRaw(sheetName) {
  if (GLOBAL_SHEET_CACHE[sheetName]) {
    return GLOBAL_SHEET_CACHE[sheetName];
  }

  const cache = CacheService.getScriptCache();
  const key = 'DS_raw_' + sheetName.replace(/[^a-zA-Z0-9]/g, '_');
  const cached = cacheGetBig_(key);
  
  if (cached) { 
    try { 
      const parsed = JSON.parse(cached);
      GLOBAL_SHEET_CACHE[sheetName] = parsed;
      return parsed; 
    } catch (e) {} 
  }
  
  const ss = getSS();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet tidak ditemukan: ' + sheetName);
  
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn(); 
  if (lastRow < 2) return { headers: [], rows: [] };
  
  const all = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const result = { headers: all[0], rows: all.slice(1) };

  // Sheet yang sedang menghitung ulang (Loading… / #N/A / #REF!) tidak disimpan ke cache
  const isTransient = v => typeof v === 'string' && (v === 'Loading...' || v === 'Loading…' || /^#(N\/A|REF!|VALUE!|ERROR!|DIV\/0!|NAME\?)/.test(v));
  const hasTransient = result.rows.some(row => row.some(isTransient));

  if (!hasTransient) {
    try {
      const jsonStr = JSON.stringify(result);
      cachePutBig_(key, jsonStr, CACHE_TTL_SECONDS);
    } catch (e) {}
  }
  
  if (!hasTransient) GLOBAL_SHEET_CACHE[sheetName] = result;
  return result;
}

/* ─── PEMBACAAN BOBOT METRIK DARI SHEET "Matriks & Bobot" ─── */
function getMetricWeights() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('METRIC_WEIGHTS_MAP');
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }

  const weights = { FM: {}, LM: {} };
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.WEIGHTS || 'Matriks & Bobot');
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const fmMetric = String(data[i][0] || '').trim();
        const fmWeight = parseFloat(data[i][1]);
        if (fmMetric) {
          weights.FM[fmMetric] = !isNaN(fmWeight) ? fmWeight : 1;
        }

        const lmMetric = String(data[i][3] || '').trim();
        const lmWeight = parseFloat(data[i][4]);
        if (lmMetric) {
          weights.LM[lmMetric] = !isNaN(lmWeight) ? lmWeight : 1;
        }
      }
    }
  } catch (e) {
    Logger.log('Gagal membaca sheet Matriks & Bobot: ' + e.message);
  }

  try { cache.put('METRIC_WEIGHTS_MAP', JSON.stringify(weights), CACHE_TTL_SECONDS); } catch (e) {}
  return weights;
}

function fmtDate(d) {
  if (!d) return '';
  if (d instanceof Date) return Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  if (typeof d === 'string' && /^\d{4}-\d{2}-\d{2}T/.test(d)) return Utilities.formatDate(new Date(d), CONFIG.TIMEZONE, 'yyyy-MM-dd');
  return String(d);
}

function fmtMonth(d) {
  if (!d) return '';
  if (d instanceof Date) return Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM');
  if (typeof d === 'string' && /^\d{4}-\d{2}-\d{2}T/.test(d)) return Utilities.formatDate(new Date(d), CONFIG.TIMEZONE, 'yyyy-MM');
  const s = String(d);
  return s.length >= 7 ? s.slice(0, 7) : s;
}

function achToStatus(ach) {
  if (ach === null || ach === undefined || ach === '' || isNaN(ach)) return 'nodata';
  let a = Number(ach);
  if (a <= 1 + 1e-9) a = a * 100; // data bisa berupa pecahan (0.99) atau persen (99)
  return a >= CONFIG.TARGET_PCT ? 'tercapai' : (a >= CONFIG.WARN_PCT ? 'waspada' : 'kritis');
}

/* ─── PERBAIKAN BACA NILAI 0 PADA METRIK ─── */

function extractMetricsByIndices(rawRow, indices, metricList) {
  const metrics = {};
  for (let i = 0; i < indices.length; i++) {
    const v = rawRow[indices[i]];
    // Jika v bernilai 0, tetap dipertahankan sebagai angka 0
    metrics[metricList[i].key] = (v === null || v === undefined || v === '' || isNaN(v)) ? null : Number(v);
  }
  return metrics;
}

function extractMetricsConsecutive(rawRow, startIdx, count, metricList) {
  const metrics = {};
  for (let i = 0; i < count; i++) {
    const v = rawRow[startIdx + i];
    // Jika v bernilai 0, tetap dipertahankan sebagai angka 0
    metrics[metricList[i].key] = (v === null || v === undefined || v === '' || isNaN(v)) ? null : Number(v);
  }
  return metrics;
}

function normalizeFMDailyRow(rawRow, segment) {
  const c = CONFIG.FM_COLUMNS.DAILY;
  return { segment, date: fmtDate(rawRow[c.DATE]), week: String(rawRow[c.WEEK]||''), month: fmtMonth(rawRow[c.MONTH]), hub: String(rawRow[c.HUB]||'').trim(), metrics: extractMetricsConsecutive(rawRow, c.METRICS_START, c.METRICS_COUNT, CONFIG.FM_METRICS) };
}
function normalizeFMWeeklyRow(rawRow) {
  const c = CONFIG.FM_COLUMNS.WEEKLY;
  const score = rawRow[c.SCORE];
  return { week: String(rawRow[c.WEEK] || ''), month: fmtMonth(rawRow[c.MONTH]), hub: String(rawRow[c.HUB] || '').trim(), score: isNaN(score) ? null : Number(score), metrics: extractMetricsConsecutive(rawRow, c.METRICS_START, c.METRICS_COUNT, CONFIG.FM_METRICS), metricScores: extractMetricsConsecutive(rawRow, c.SCORE_START, c.METRICS_COUNT, CONFIG.FM_METRICS) };
}
function normalizeFMMonthlyRow(rawRow) {
  const c = CONFIG.FM_COLUMNS.MONTHLY;
  const score = rawRow[c.SCORE];
  return { month: fmtMonth(rawRow[c.MONTH]), hub: String(rawRow[c.HUB]||'').trim(), score: isNaN(score)?null:Number(score), metrics: extractMetricsConsecutive(rawRow, c.METRICS_START, c.METRICS_COUNT, CONFIG.FM_METRICS) };
}

function normalizeLMDailyRow(rawRow) {
  const c = CONFIG.LM_COLUMNS.DAILY;
  return { segment: 'regular', date: fmtDate(rawRow[c.DATE]), week: String(rawRow[c.WEEK]||''), month: fmtMonth(rawRow[c.MONTH]), hub: String(rawRow[c.HUB]||'').trim(), metrics: extractMetricsByIndices(rawRow, c.METRIC_INDICES, CONFIG.LM_METRICS) };
}
function normalizeLMWeeklyRow(rawRow) {
  const c = CONFIG.LM_COLUMNS.WEEKLY;
  const score = rawRow[c.SCORE];
  return { week: String(rawRow[c.WEEK]||''), month: '', hub: String(rawRow[c.HUB]||'').trim(), score: isNaN(score)?null:Number(score), metrics: extractMetricsByIndices(rawRow, c.METRIC_INDICES, CONFIG.LM_METRICS), metricScores: extractMetricsByIndices(rawRow, c.SCORE_INDICES, CONFIG.LM_METRICS) };
}

/* Normalisasi baris sheet dilakukan sekali per eksekusi (getHubList/getMonthList/getWeekList/getDateList
   sebelumnya masing-masing memetakan ulang seluruh sheet harian). Jangan memodifikasi objek baris yang dikembalikan. */
const ROWS_MEMO_ = {};
function memoRows_(k, fn) { if (!ROWS_MEMO_[k]) ROWS_MEMO_[k] = fn(); return ROWS_MEMO_[k]; }

function getFMRegular() { return memoRows_('getFMRegular', function () { return readSheetRaw(CONFIG.SHEETS.FM_REGULAR).rows.map(r => normalizeFMDailyRow(r, 'regular')).filter(r => r.hub && r.date); }); }
function getFMHigh() { return memoRows_('getFMHigh', function () { return readSheetRaw(CONFIG.SHEETS.FM_HIGH).rows.map(r => normalizeFMDailyRow(r, 'high')).filter(r => r.hub && r.date); }); }
function getFMWeekly() { return memoRows_('getFMWeekly', function () { return readSheetRaw(CONFIG.SHEETS.FM_WEEKLY).rows.map(normalizeFMWeeklyRow).filter(r => r.hub); }); }
function getFMMonthly() { return memoRows_('getFMMonthly', function () { return readSheetRaw(CONFIG.SHEETS.FM_MONTHLY).rows.map(normalizeFMMonthlyRow).filter(r => r.hub); }); }

function getLMDaily() { return memoRows_('getLMDaily', function () { return readSheetRaw(CONFIG.SHEETS.LM_BACKUP).rows.map(normalizeLMDailyRow).filter(r => r.hub && r.date); }); }
function getLMWeekly() { return memoRows_('getLMWeekly', function () { return readSheetRaw(CONFIG.SHEETS.LM_BACKUP).rows.map(normalizeLMWeeklyRow).filter(r => r.hub && r.week); }); }

function cleanHubName(name) {
  if (!name) return '';
  return String(name).toLowerCase().replace(/first mile hub/g, '').replace(/hub/g, '').replace(/[^a-z0-9]/g, '').trim();
}

function filterByHub(rows, hubFilter) { 
  if (!hubFilter || hubFilter === 'Semua') return rows; 
  const cleanTarget = cleanHubName(hubFilter);
  return rows.filter(r => cleanHubName(r.hub) === cleanTarget);
}

function getHubList(mode) {
  try {
    const set = new Set();
    if (mode === 'LM') getLMDaily().forEach(r => set.add(r.hub));
    else { getFMRegular().forEach(r => set.add(r.hub)); getFMHigh().forEach(r => set.add(r.hub)); }
    return [...set].filter(h => h && /[a-zA-Z]/.test(h)).sort();
  } catch (e) { return []; }
}

function getMonthList(mode) {
  try {
    const set = new Set();
    const rows = mode === 'LM' ? getLMDaily() : getFMRegular().concat(getFMHigh());   // High Days ikut, agar bulan/minggu yang baru ada di High tetap muncul
    rows.forEach(r => r.month && set.add(r.month));
    return [...set].sort().reverse();
  } catch (e) { return []; }
}

function getWeekList(mode) {
  try {
    const map = new Map();
    const rows = mode === 'LM' ? getLMDaily() : getFMRegular().concat(getFMHigh());   // sebelumnya hanya Reguler Days → minggu yang baru ada di High Days tidak muncul
    rows.forEach(r => { if (r.week && (r.month || !map.has(r.week))) map.set(r.week, r.month || ''); });
    return [...map.entries()].map(([value, month]) => ({ value, month })).sort((a, b) => b.month.localeCompare(a.month) || b.value.localeCompare(a.value));
  } catch (e) { return []; }
}

function getDateList(mode) {
  const currentMode = mode || 'FM';
  const dateMap = new Map();
  const rows = currentMode === 'LM' ? getLMDaily() : (getFMRegular() || []).concat(getFMHigh() || []);

  rows.forEach(r => {
    if (r.date) {
      let dateStr = r.date instanceof Date ? Utilities.formatDate(r.date, CONFIG.TIMEZONE, 'yyyy-MM-dd') : String(r.date).trim();
      if (dateStr && !dateMap.has(dateStr)) {
        dateMap.set(dateStr, { value: dateStr, week: String(r.week || '').trim(), month: String(r.month || '').trim() });
      }
    }
  });

  return Array.from(dateMap.values()).sort((a, b) => a.value.localeCompare(b.value));
}

function filterByDateWeekMonth(rows, date, week, month) {
  if (!rows || !rows.length) return rows;
  if (date) return rows.filter(r => r.date === date);
  if (week) return rows.filter(r => r.week === week);
  if (month) return rows.filter(r => r.month === month);
  return rows;
}

function filterByMonth(rows, month) { return !month ? rows : rows.filter(r => r.month === month); }

/* ─── CACHE BESAR: data dipecah per 30.000 karakter agar sheet >100KB tetap ter-cache ─── */
function hashKey_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s).map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}
function cachePutBig_(key, str, ttl) {
  if (!str || str.length > 1500000) return;
  const cache = CacheService.getScriptCache(), size = 30000, n = Math.ceil(str.length / size), o = {};
  for (let i = 0; i < n; i++) o[key + '_' + i] = str.substr(i * size, size);
  o[key + '_n'] = String(n);
  try { cache.putAll(o, ttl || CACHE_TTL_SECONDS); } catch (e) {}
}
function cacheGetBig_(key) {
  const cache = CacheService.getScriptCache();
  const n = parseInt(cache.get(key + '_n'), 10);
  if (!n) return null;
  const keys = []; for (let i = 0; i < n; i++) keys.push(key + '_' + i);
  const m = cache.getAll(keys); let out = '';
  for (let i = 0; i < n; i++) { const p = m[keys[i]]; if (p === undefined || p === null) return null; out += p; }
  return out;
}
function clearBigCache_(cache, bases) {
  bases.forEach(b => {
    const n = parseInt(cache.get(b + '_n'), 10) || 0, ks = [b + '_n'];
    for (let i = 0; i < n; i++) ks.push(b + '_' + i);
    cache.removeAll(ks);
  });
}