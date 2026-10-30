/************************************************************
 * CODE.GS — Unified FM & LM Backend API · JTT 07
 ************************************************************/

/* Favicon tab browser (logo JTT07), di-host di GitHub (URL Raw, PNG langsung).
   Google Drive tidak bisa dipakai untuk setFaviconUrl. Kosongkan = ikon bawaan Apps Script. */
const FAVICON_URL = 'https://raw.githubusercontent.com/andikapradana-gif/jtt07-assets/main/favicon-JTT07-64.png';

function doGet(e) {
  const userEmail = Session.getActiveUser().getEmail();
  const access = getUserAccess(userEmail, 'FM');

  if (!access.isAllowed) {
    return HtmlService.createHtmlOutput(
      `<div style="font-family: 'Inter', sans-serif; padding: 60px 20px; text-align: center; background: #f5f5f7; height: 100vh;">
         <div style="background: white; max-width: 400px; margin: 0 auto; padding: 40px; border-radius: 20px; box-shadow: 0 10px 30px rgba(0,0,0,0.05);">
           <h1 style="color: #ef4444; font-size: 24px; margin-bottom: 16px;">Akses Ditolak</h1>
           <p style="color: #64748b; font-size: 14px; line-height: 1.6;">Maaf, email <b>${userEmail || 'Anda'}</b> tidak terdaftar dalam sistem <b>Access Web</b>.</p>
           <p style="color: #94a3b8; font-size: 12px; margin-top: 24px;">Silakan hubungi Administrator JTT-07.</p>
         </div>
       </div>`
    );
  }

  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.initState = sanitizeInitState_(e);
  const out = tpl
    .evaluate().setTitle('Dashboard - JTT 07')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  if (FAVICON_URL) {
    try { out.setFaviconUrl(FAVICON_URL); }
    catch (err) { console.warn('Favicon gagal dipasang: ' + err); }
  }
  return out;
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }
function ok(data, meta) { return { success: true, data: data, error: null, meta: meta || {} }; }
function fail(errorMsg) { return { success: false, data: null, error: String(errorMsg), meta: {} }; }

const API_CACHE_TTL = 600;

/* Kunci cache memuat cakupan akses pengguna, sehingga data hub yang dibatasi tidak bocor ke pengguna lain */
function accessScope_() {
  const e = Session.getActiveUser().getEmail();
  const f = getUserAccess(e, 'FM'), l = getUserAccess(e, 'LM');
  return [f.role || 'x', (f.hubAccess || []).join(','), (l.hubAccess || []).join(',')].join('|');
}

/* ═══ SKOR RESMI = kolom SCORE di FM Blended Weekly / FM Blended Monthly / LM [Backup NEW] ═══
   Satu sumber untuk Ringkasan, Trend, selisih WoW, dan kartu landing. */
/* Pecahan (0.99 / 1) → persen. Toleransi 1e-9: sel Sheets sering menghasilkan 1.0000000000000002 untuk nilai 100%,
   yang sebelumnya dibaca sebagai "1%" dan membuat hub ber-skor sempurna terbuang dari rata-rata. */
function isFraction_(n) { return n <= 1 + 1e-9; }
function scoreToPct_(x) { const n = Number(x); return isFraction_(n) ? n * 100 : n; }

/* Baris dianggap punya data bila minimal satu metrik terisi non-nol (rumus SCORE bisa sudah ada walau metrik belum masuk) */
function rowHasMetricData_(r) {
  return !!(r && r.metrics && Object.keys(r.metrics).some(k => {
    const v = r.metrics[k];
    return v !== null && v !== undefined && !isNaN(v) && Number(v) !== 0;
  }));
}

/* Skor jauh di bawah rata-rata metriknya = hasil rumus setengah-jadi */
function rowScoreSuspect_(r) {
  const vals = Object.keys(r.metrics || {}).map(k => r.metrics[k])
    .filter(v => v !== null && v !== undefined && !isNaN(v)).map(scoreToPct_);
  if (vals.length < 3) return false;
  const mean = vals.reduce((a, b) => a + b, 0) / vals.length;
  return mean >= 50 && scoreToPct_(r.score) < mean - 30;
}

function rowScoreUsable_(r) {
  return r && r.score !== null && r.score !== undefined && r.score !== '' && !isNaN(r.score)
    && rowHasMetricData_(r) && !rowScoreSuspect_(r);
}

/* Rata-rata SCORE: per hub dulu (jika ada baris ganda), lalu antar hub. Hasil dalam persen. */
function avgHubScore_(rows) {
  const byHub = {}, seen = {};
  (rows || []).forEach(r => {
    seen[r.hub] = true;
    if (!rowScoreUsable_(r)) return;
    (byHub[r.hub] = byHub[r.hub] || []).push(scoreToPct_(r.score));
  });
  const list = Object.keys(byHub).map(h => ({ hub: h, score: byHub[h].reduce((a, b) => a + b, 0) / byHub[h].length }));
  const excludedHubs = Object.keys(seen).filter(h => !byHub[h]).sort();
  return {
    avg: list.length ? list.reduce((a, b) => a + b.score, 0) / list.length : null,
    hubs: list.length,
    excluded: excludedHubs.length,
    excludedHubs: excludedHubs,
    list: list
  };
}

function withApiCache(cacheKey, fn, isCacheable) {
  const cache = CacheService.getScriptCache();
  let key = 'API_V20_' + (cache.get('API_VER') || '0') + '_' + hashKey_(accessScope_()) + '_' + cacheKey;
  if (key.length > 150) key = 'API_H_' + hashKey_(key);
  const cached = cacheGetBig_(key);
  if (cached) { try { return JSON.parse(cached); } catch (e) {} }

  const result = fn();
  // Hasil gagal / tidak lolos validasi TIDAK disimpan, agar data setengah-jadi tidak "terkunci" di cache
  let cacheable = !(result && result.success === false);
  if (cacheable && typeof isCacheable === 'function') { try { cacheable = !!isCacheable(result); } catch (e) { cacheable = false; } }
  if (cacheable) { try { cachePutBig_(key, JSON.stringify(result), API_CACHE_TTL); } catch (e) {} }
  return result;
}

function apiClearCache() {
  try {
    const cache = CacheService.getScriptCache();
    clearBigCache_(cache, ['DS_raw_FM_Reguler_Days', 'DS_raw_FM_High_Days', 'DS_raw_FM_Blended_Weekly', 'DS_raw_FM_Blended_Monthly', 'DS_raw_LM__Backup_NEW_']);
    cache.removeAll(['HYPERCARE_SAVED_INPUTS', 'METRIC_WEIGHTS_MAP', 'ACCESS_ROWS']);
    cache.put('API_VER', String(Date.now()), 21600); // membuat semua cache API lama tidak terpakai
    return ok({ message: 'Cache server berhasil dibersihkan!' });
  } catch (e) { return fail(e.message); }
}

/* Pengguna aktif: heartbeat tiap 15 detik dari browser; dianggap aktif bila terakhir terlihat < 50 detik.
   Penulisan daftar memakai kunci (lock) agar dua pengguna yang masuk bersamaan tidak saling menimpa. */
const ACTIVE_WINDOW_MS = 50000;
function activeMapUpdate_(mutator) {
  const cache = CacheService.getScriptCache();
  let lock = null, locked = false;
  try { lock = LockService.getScriptLock(); locked = lock.tryLock(3000); } catch (e) {}
  try {
    const map = JSON.parse(cache.get('ONLINE_USERS_MAP') || '{}');
    mutator(map);
    const now = Date.now();
    Object.keys(map).forEach(u => { if (now - map[u] > ACTIVE_WINDOW_MS) delete map[u]; });
    cache.put('ONLINE_USERS_MAP', JSON.stringify(map), 300);
    return map;
  } finally { if (locked) { try { lock.releaseLock(); } catch (e) {} } }
}

function apiSendHeartbeat() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) return ok({ status: 'ignored', users: [] });
    const map = activeMapUpdate_(m => { m[email] = Date.now(); });
    const users = Object.keys(map);
    return ok({ status: 'active', onlineCount: users.length, users: users });   // daftar ikut dikirim: satu panggilan cukup
  } catch (e) { return fail(e.message); }
}

/* Dipanggil saat pengguna menutup tab / menyembunyikan tab / kembali ke landing: nama langsung dihapus dari daftar aktif */
function apiLeaveActive() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) return ok({ status: 'ignored' });
    activeMapUpdate_(m => { delete m[email]; });
    return ok({ status: 'left' });
  } catch (e) { return fail(e.message); }
}

function apiGetActiveUsers() {
  try {
    const map = JSON.parse(CacheService.getScriptCache().get('ONLINE_USERS_MAP') || '{}');
    const now = Date.now();
    return ok({ users: Object.keys(map).filter(u => now - map[u] <= ACTIVE_WINDOW_MS) });
  } catch (e) { return fail(e.message); }
}

function buildHypercareKey(hub, week, metricKey) {
  const cleanH = cleanHubName(hub);
  const cleanW = String(week || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const cleanM = String(metricKey || '').toLowerCase().trim();
  return `${cleanH}_${cleanW}_${cleanM}`;
}

function saveHypercareInputUnsafe_(hub, week, metricKey, rootCause, actionPlan, status, dueDate, pic, whyChain) {
  try {
    const ss = getSS();
    let sheet = ss.getSheetByName('Hypercare_Input');
    if (!sheet) {
      sheet = ss.insertSheet('Hypercare_Input');
      sheet.appendRow(['Key_ID', 'Hub', 'Week', 'Metric_Key', 'Root_Cause', 'Action_Plan', 'Updated_By', 'Updated_At', 'Status', 'Due_Date', 'PIC', 'Why_Chain']);
      sheet.getRange("A1:L1").setFontWeight("bold").setBackground("#f3f3f3");
      sheet.setFrozenRows(1);
    }

    const email = Session.getActiveUser().getEmail() || 'Anonymous';
    const timestamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    sheet.getRange(1, 9, 1, 4).setValues([['Status', 'Due_Date', 'PIC', 'Why_Chain']]);
    const keyId = buildHypercareKey(hub, week, metricKey);

    const cleanH = cleanHubName(hub);
    const cleanW = String(week || '').toLowerCase().replace(/[^a-z0-9]/g, '');
    const cleanM = String(metricKey || '').toLowerCase().trim();

    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      const rowKey = String(data[i][0]).toLowerCase().trim();
      const rowH = cleanHubName(data[i][1]);
      const rowW = String(data[i][2]).toLowerCase().replace(/[^a-z0-9]/g, '');
      const rowM = String(data[i][3]).toLowerCase().trim();

      if (rowKey === keyId || (rowH === cleanH && rowW === cleanW && rowM === cleanM)) { 
        rowIndex = i + 1; 
        break; 
      }
    }

    if (rowIndex > 0) {
      const keepWhy = (whyChain === null || whyChain === undefined) ? (data[rowIndex - 1][11] || '') : whyChain;   // null = jangan ubah
      sheet.getRange(rowIndex, 1, 1, 12).setValues([[keyId, hub, week, metricKey, rootCause, actionPlan, email, timestamp, status || 'Open', dueDate || '', pic || '', keepWhy]]);
    } else {
      sheet.appendRow([keyId, hub, week, metricKey, rootCause, actionPlan, email, timestamp, status || 'Open', dueDate || '', pic || '', whyChain || '']);
    }

    CacheService.getScriptCache().remove('HYPERCARE_SAVED_INPUTS');
    return ok({ message: 'Root Cause & Action Plan berhasil disimpan!' });
  } catch (e) { 
    return fail(e.message); 
  }
}

function getHypercareSavedInputs() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('HYPERCARE_SAVED_INPUTS');
  if (cached) { try { return JSON.parse(cached); } catch (e) {} }

  const map = {};
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName('Hypercare_Input');
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0]) {
          const rawKey = String(data[i][0]).toLowerCase().trim();
          const hub = data[i][1];
          const week = data[i][2];
          const metricKey = data[i][3];

          const valObj = {
            rootCause: data[i][4] || '', 
            actionPlan: data[i][5] || '',
            updatedBy: data[i][6] || '', status: data[i][8] || 'Open', dueDate: data[i][9] ? (data[i][9] instanceof Date ? Utilities.formatDate(data[i][9], CONFIG.TIMEZONE, 'yyyy-MM-dd') : String(data[i][9])) : '', pic: data[i][10] || '', whyChain: data[i][11] || '', 
            temuan: hcvParse_(data[i][12]), val: hcvParse_(data[i][13]),
            updatedAt: data[i][7] ? Utilities.formatDate(new Date(data[i][7]), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm') : ''
          };

          map[rawKey] = valObj;
          if (hub && week && metricKey) {
            const exactKey = buildHypercareKey(hub, week, metricKey);
            map[exactKey] = valObj;
          }
        }
      }
    }
  } catch (e) {}

  try { cache.put('HYPERCARE_SAVED_INPUTS', JSON.stringify(map), 300); } catch (e) {}
  return map;
}

/* Hasil akses dihitung sekali per eksekusi: accessScope_() + tiap api* memanggilnya berkali-kali */
const ACCESS_MEMO_ = {};
function getUserAccess(email, mode) {
  const k = String(email || '').toLowerCase() + '|' + (mode || 'FM');
  if (!ACCESS_MEMO_[k]) ACCESS_MEMO_[k] = getUserAccessRaw_(email, mode);
  return ACCESS_MEMO_[k];
}
function getUserAccessRaw_(email, mode) {
  try {
    if (!email) return { isAllowed: false };
    const data = getAccessRows_();
    if (!data) return { isAllowed: false, error: 'Sheet Access Web tidak ditemukan' };
    const currentMode = mode || 'FM';

    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][1] || '').toLowerCase().trim();
      if (rowEmail === String(email).toLowerCase().trim()) {
        const role = String(data[i][2] || '').trim();
        let allowedHubs = [];

        if (role === 'City Lead' || role === 'PIC') allowedHubs = ['Semua'];
        else {
          if (currentMode === 'LM') {
            const lm1 = String(data[i][4] || '').trim();
            const lm2 = String(data[i][5] || '').trim();
            if (lm1 && lm1 !== '-' && lm1.toLowerCase() !== 'null') allowedHubs.push(lm1);
            if (lm2 && lm2 !== '-' && lm2.toLowerCase() !== 'null') allowedHubs.push(lm2);
          } else {
            const fmRaw = String(data[i][3] || '').trim();
            if (fmRaw && fmRaw !== '-' && fmRaw.toLowerCase() !== 'null') {
              allowedHubs = fmRaw.split(',').map(h => h.trim()).filter(Boolean);
            }
          }
        }

        return { isAllowed: true, email: email, role: role, hubAccess: allowedHubs.length > 0 ? allowedHubs : ['Semua'] };
      }
    }
    return { isAllowed: false };
  } catch (e) { return { isAllowed: false, error: e.message }; }
}

/* Opsi filter di-cache di server (sebelumnya dihitung ulang dari sheet harian setiap dashboard dibuka) */
function apiGetFilterOptions(mode) {
  const m = mode === 'LM' ? 'LM' : 'FM';
  return withApiCache('filteropts_v2_' + m, function () { return apiGetFilterOptionsRaw_(m); });
}
function apiGetFilterOptionsRaw_(mode) {
  try {
    const currentMode = mode || 'FM';
    const access = getUserAccess(Session.getActiveUser().getEmail(), currentMode);
    let allHubsFromSheet = getHubList(currentMode);
    let matchedHubs = [];

    if (access.hubAccess && !access.hubAccess.includes('Semua')) {
      const normalizeStrict = (str) => String(str || '').toLowerCase().replace(/[^a-z0-9]/g, '').trim();
      const userStrictList = access.hubAccess.map(normalizeStrict);
      matchedHubs = allHubsFromSheet.filter(sheetHub => userStrictList.includes(normalizeStrict(sheetHub)));
      if (matchedHubs.length === 0) matchedHubs = access.hubAccess;
    } else matchedHubs = allHubsFromSheet;

    return ok({ hubs: matchedHubs, months: getMonthList(currentMode), weeks: getWeekList(currentMode), dates: getDateList(currentMode), user: access, target: CONFIG.TARGET_PCT, warn: CONFIG.WARN_PCT, latestWeek: latestWeekLabel_(currentMode) });
  } catch (e) { return fail(e.message); }
}

/* ─── API RINGKASAN DASHBOARD ─── */
function apiGetRingkasan(mode, hubFilter, date, week, month) {
  const currentMode = mode || 'FM';
  const userEmail = Session.getActiveUser().getEmail();
  const cacheKey = `ringkasan_v40_${currentMode}_${hubFilter || 'all'}_${date || '-'}_${week || '-'}_${month || '-'}`;
  return withApiCache(cacheKey, function() {
    try {
      const access = getUserAccess(userEmail, currentMode);
      let finalHubFilter = hubFilter;
      
      if (access.hubAccess && !access.hubAccess.includes('Semua') && access.hubAccess.length > 0) {
        if (!finalHubFilter || finalHubFilter === 'Semua' || !access.hubAccess.includes(finalHubFilter)) {
          finalHubFilter = access.hubAccess[0];
        }
      }

      const allWeights = getMetricWeights();
      const modeWeights = allWeights[currentMode] || {};
      const metricsList = currentMode === 'LM' ? CONFIG.LM_METRICS : CONFIG.FM_METRICS;

      const calcAvg = (rows, key) => {
        let total = 0, count = 0;
        rows.forEach(r => {
          const v = r.metrics ? r.metrics[key] : null;
          if (v !== null && v !== undefined && !isNaN(v)) {
            let numVal = Number(v);
            if (isFraction_(numVal)) numVal = numVal * 100;
            total += numVal;
            count++;
          }
        });
        return count > 0 ? total / count : null;
      };

      const memo = {};
      const memoBase = (k, f) => (memo[k] = memo[k] || f());
      const findWorstHubForMetricStrict = (metricKey, dateVal, weekVal, monthVal) => {
        let rawDetail = [];
        if (currentMode === 'LM') {
          rawDetail = memoBase('LMW', getLMWeekly);
        } else {
          if (monthVal && !weekVal && !dateVal) rawDetail = memoBase('FMM', getFMMonthly);
          else rawDetail = memoBase('FMW', getFMWeekly);
        }

        if (access.hubAccess && !access.hubAccess.includes('Semua') && access.hubAccess.length > 0) {
          rawDetail = rawDetail.filter(r => access.hubAccess.includes(r.hub));
        } else if (finalHubFilter && finalHubFilter !== 'Semua') {
          rawDetail = filterByHub(rawDetail, finalHubFilter);
        }

        if (weekVal) {
          rawDetail = rawDetail.filter(r => String(r.week) === String(weekVal));
          if (!rawDetail.length && provisional) rawDetail = blendRows;   // minggu berjalan: pakai baris harian
        } else if (monthVal && currentMode === 'FM') {
          rawDetail = rawDetail.filter(r => String(r.month) === String(monthVal));
        }

        const hubAgg = {};
        rawDetail.forEach(r => {
          const v = r.metrics ? r.metrics[metricKey] : null;
          if (v !== null && v !== undefined && !isNaN(v)) {
            let numVal = isFraction_(Number(v)) ? Number(v) * 100 : Number(v);
            if (!hubAgg[r.hub]) hubAgg[r.hub] = { sum: 0, count: 0 };
            hubAgg[r.hub].sum += numVal;
            hubAgg[r.hub].count += 1;
          }
        });

        let worstHubName = (finalHubFilter && finalHubFilter !== 'Semua') ? finalHubFilter : '—';
        let minScore = Infinity;

        Object.keys(hubAgg).forEach(h => {
          const avg = hubAgg[h].sum / hubAgg[h].count;
          if (avg < minScore) {
            minScore = avg;
            worstHubName = h;
          }
        });

        return {
          hub: worstHubName,
          val: minScore !== Infinity ? minScore : null
        };
      };

      let blendRows = [];
      let regRows = [];
      let highRows = [];
      let provisional = false;   // true = minggu belum ada di FM Blended Weekly → dihitung sementara dari data harian

      if (currentMode === 'FM') {
        if (month && !week && !date) {
          regRows = filterByMonth(filterByHub(getFMRegular(), finalHubFilter), month);
          highRows = filterByMonth(filterByHub(getFMHigh(), finalHubFilter), month);
          blendRows = filterByMonth(filterByHub(getFMMonthly(), finalHubFilter), month);
        } else if (week && !date) {
          regRows = filterByHub(getFMRegular(), finalHubFilter).filter(r => String(r.week) === String(week));
          highRows = filterByHub(getFMHigh(), finalHubFilter).filter(r => String(r.week) === String(week));
          blendRows = filterByHub(getFMWeekly(), finalHubFilter).filter(r => String(r.week) === String(week));
        } else {
          regRows = filterByDateWeekMonth(filterByHub(getFMRegular(), finalHubFilter), date, week, month);
          highRows = filterByDateWeekMonth(filterByHub(getFMHigh(), finalHubFilter), date, week, month);
          blendRows = regRows.concat(highRows);
        }
      } else {
        if (week && !date) {
          blendRows = filterByHub(getLMWeekly(), finalHubFilter).filter(r => String(r.week) === String(week));
        } else {
          blendRows = filterByDateWeekMonth(filterByHub(getLMDaily(), finalHubFilter), date, week, month);
        }
        regRows = blendRows;
      }

      if (currentMode === 'FM' && week && !date && !blendRows.length && (regRows.length || highRows.length)) {
        blendRows = regRows.concat(highRows);
        provisional = true;
      }

      const criticalAnalysis = [];
      let totalWeightedScore = 0, totalWeight = 0;

      const summary = metricsList.map(m => {
        const regVal = calcAvg(regRows, m.key);
        const highVal = calcAvg(highRows, m.key);
        const blendVal = calcAvg(blendRows, m.key);

        const weight = modeWeights[m.key] || modeWeights[m.label] || 1;
        if (blendVal !== null) {
          totalWeightedScore += blendVal * weight;
          totalWeight += weight;
        }

        if (blendVal !== null && blendVal < CONFIG.TARGET_PCT) {
          const gap = parseFloat((100 - blendVal).toFixed(2));
          const weightedImpact = parseFloat((gap * weight).toFixed(2));
          const worstInfo = findWorstHubForMetricStrict(m.key, date, week, month);

          criticalAnalysis.push({
            metricKey: m.key, 
            metricLabel: m.label, 
            blendedScore: blendVal, 
            gap: gap,
            weight: weight, 
            weightedImpact: weightedImpact, 
            worstHub: worstInfo.hub, 
            worstVal: worstInfo.val !== null ? worstInfo.val : blendVal,
            clearSuggestion: `Tindakan: ${CONFIG.DEFAULT_ACTIONS[m.key] || 'Lakukan audit operasional.'}`
          });
        }

        return {
          key: m.key, label: m.label, target: 100, weight: weight,
          regular: regVal, high: highVal, blended: blendVal,
          statusRegular: achToStatus(regVal !== null ? regVal / 100 : null),
          statusHigh: achToStatus(highVal !== null ? highVal / 100 : null),
          statusBlended: achToStatus(blendVal !== null ? blendVal / 100 : null)
        };
      });

      // Skor resmi = kolom SCORE (tersedia pada tampilan mingguan & bulanan). Tampilan harian tidak punya SCORE → rata-rata metrik berbobot.
      const scoreView = !provisional && (currentMode === 'FM' ? ((month && !week && !date) || (week && !date)) : (week && !date));
      const scoreInfo = scoreView ? avgHubScore_(blendRows) : null;
      const hasScore = !!(scoreInfo && scoreInfo.avg !== null);
      const overallScore = hasScore ? scoreInfo.avg : (totalWeight > 0 ? totalWeightedScore / totalWeight : null);
      const overallSource = hasScore ? 'score' : 'metrics';
      criticalAnalysis.sort((a, b) => b.weightedImpact - a.weightedImpact);

      // Skor dianggap TIDAK LENGKAP bila ada hub yang barisnya ada tetapi belum punya skor valid (rumus masih menghitung / data belum masuk).
      // Hasil seperti ini tidak boleh di-cache agar tidak "terkunci" selama TTL.
      const scoreIncomplete = !!(scoreInfo && (scoreInfo.excluded > 0 || (scoreInfo.avg === null && blendRows.length > 0)));

      return ok({ mode: currentMode, summary, criticalAnalysis, overallScore, overallSource, scoreHubs: hasScore ? scoreInfo.hubs : null,
        scoreTotal: scoreInfo ? scoreInfo.hubs + scoreInfo.excluded : null, scoreExcluded: scoreInfo ? scoreInfo.excludedHubs : null, scoreIncomplete: scoreIncomplete,
        effectiveMonth: month, provisional: provisional, hub: finalHubFilter || 'Semua Hub', wow: computeWow_(currentMode, finalHubFilter, week) });
    } catch (e) { 
      return fail(e.message); 
    }
  }, function (res) { return !(res && res.success && res.data && res.data.scoreIncomplete); });
}

/* ─── API BARU: DAILY REPORT MATRIX ─── */
function apiGetDailyReportMatrix(mode, segment, hubFilter, month, week) {
  const currentMode = mode || 'FM';
  const seg = segment || 'regular';
  const cacheKey = `daily_matrix_${currentMode}_${seg}_${hubFilter || 'all'}_${month || '-'}_${week || '-'}`;

  return withApiCache(cacheKey, function() {
    try {
      const access = getUserAccess(Session.getActiveUser().getEmail(), currentMode);
      let finalHubFilter = hubFilter;
      if (access.hubAccess && !access.hubAccess.includes('Semua') && access.hubAccess.length > 0) {
        if (!finalHubFilter || !access.hubAccess.includes(finalHubFilter)) finalHubFilter = access.hubAccess[0];
      }

      let rows = currentMode === 'LM' 
        ? getLMDaily() 
        : (seg === 'high' ? getFMHigh() : getFMRegular());

      rows = filterByHub(rows, finalHubFilter);
      rows = filterByDateWeekMonth(rows, null, week, month);

      const metricsList = currentMode === 'LM' ? CONFIG.LM_METRICS : CONFIG.FM_METRICS;
      const allDates = [...new Set(rows.map(r => r.date).filter(Boolean))].sort();

      const resultRows = metricsList.map(m => {
        let totalSum = 0;
        let validCount = 0;
        let targetAchievedDays = 0;
        const datesData = {};

        allDates.forEach(d => {
          const dateRows = rows.filter(r => r.date === d);
          const vals = dateRows.map(r => r.metrics ? r.metrics[m.key] : null).filter(v => v !== null && !isNaN(v));

          if (vals.length > 0) {
            let dayAvg = (vals.reduce((a, b) => a + b, 0) / vals.length);
            if (isFraction_(dayAvg)) dayAvg = dayAvg * 100;
            datesData[d] = dayAvg;

            totalSum += dayAvg;
            validCount++;
            if (dayAvg >= CONFIG.TARGET_PCT) targetAchievedDays++;
          } else {
            datesData[d] = null;
          }
        });

        const overallAvg = validCount > 0 ? (totalSum / validCount) : null;

        return {
          metricKey: m.key,
          metricLabel: m.label,
          avg: overallAvg,
          targetCount: targetAchievedDays,
          datesData: datesData
        };
      });

      return ok({
        mode: currentMode,
        segment: seg,
        dates: allDates,
        rows: resultRows,
        hubLabel: finalHubFilter || 'Semua Hub'
      });
    } catch (e) {
      return fail(e.message);
    }
  });
}

function apiGetHarian(mode, metricKey, segment, hubFilter, month, week) {
  const currentMode = mode || 'FM';
  const seg = segment || 'blended';
  const cacheKey = `harian_${currentMode}_${metricKey}_${seg}_${hubFilter || 'all'}_${month || '-'}_${week || '-'}`;
  return withApiCache(cacheKey, function() {
    try {
      const access = getUserAccess(Session.getActiveUser().getEmail(), currentMode);
      let finalHubFilter = hubFilter;
      if (access.hubAccess && !access.hubAccess.includes('Semua') && access.hubAccess.length > 0) {
        if (!finalHubFilter || !access.hubAccess.includes(finalHubFilter)) finalHubFilter = access.hubAccess[0];
      }

      let rows = currentMode === 'LM' ? getLMDaily() : (seg === 'high' ? getFMHigh() : (seg === 'regular' ? getFMRegular() : (getFMRegular() || []).concat(getFMHigh() || [])));
      rows = filterByHub(rows, finalHubFilter);
      rows = filterByDateWeekMonth(rows, null, week, month);
      
      const metricsList = currentMode === 'LM' ? CONFIG.LM_METRICS : CONFIG.FM_METRICS;
      const metric = metricsList.find(m => m.key === metricKey) || metricsList[0];
      const allDates = [...new Set(rows.map(r => r.date).filter(Boolean))].sort();
      const hubs = [...new Set(rows.filter(r => allDates.includes(r.date)).map(r => r.hub))].sort();
      
      const matrix = hubs.map(hub => {
        const row = { hub };
        allDates.forEach(d => {
          const found = rows.filter(r => r.hub === hub && r.date === d);
          if (!found.length) { row[d] = null; return; }
          const vals = found.map(r => r.metrics[metric.key]).filter(v => v !== null && !isNaN(v));
          row[d] = vals.length ? (vals.reduce((a, b) => a + b, 0) / vals.length) * 100 : null;
        });
        return row;
      });

      return ok({ hubs, dates: allDates, matrix, metricLabel: metric.label, segment: seg, mode: currentMode });
    } catch (e) { return fail(e.message); }
  });
}

function apiGetTopLow(mode, metricKey, topN, lowN, segment, hubFilter, date, week, month) {
  const currentMode = mode || 'FM';
  const seg = segment || 'blended';
  const cacheKey = `toplow_${currentMode}_${metricKey}_${seg}_${hubFilter || 'all'}_${date || '-'}_${week || '-'}_${month || '-'}`;
  return withApiCache(cacheKey, function() {
    try {
      const access = getUserAccess(Session.getActiveUser().getEmail(), currentMode);
      let finalHubFilter = hubFilter;
      if (access.hubAccess && !access.hubAccess.includes('Semua') && access.hubAccess.length > 0) {
        if (!finalHubFilter || !access.hubAccess.includes(finalHubFilter)) finalHubFilter = access.hubAccess[0];
      }

      /* Sumber data disamakan dengan halaman Mingguan / Bulanan: bila periode = minggu (atau bulan, FM blended),
         pakai nilai sheet mingguan/bulanan, BUKAN rata-rata nilai harian (rasio/produktivitas tidak bisa dirata-ratakan per hari). */
      let rows, usedWeekly = false;
      if (!date && week && (currentMode === 'LM' || seg === 'blended')) {
        rows = (currentMode === 'LM' ? getLMWeekly() : getFMWeekly()).filter(r => String(r.week) === String(week));
        usedWeekly = true;
      } else if (!date && !week && month && currentMode === 'FM' && seg === 'blended') {
        rows = getFMMonthly().filter(r => r.month === month);
        usedWeekly = true;
      } else {
        rows = currentMode === 'LM' ? getLMDaily() : (seg === 'high' ? getFMHigh() : (seg === 'regular' ? getFMRegular() : (getFMRegular() || []).concat(getFMHigh() || [])));
        rows = filterByDateWeekMonth(rows, date, week, month);
      }
      rows = filterByHub(rows, finalHubFilter);

      const metricsList = currentMode === 'LM' ? CONFIG.LM_METRICS : CONFIG.FM_METRICS;
      const metric = metricsList.find(m => m.key === metricKey) || metricsList[0];
      const grouped = {};
      rows.forEach(r => {
        const v = r.metrics[metric.key];
        if (v === null || isNaN(v)) return;
        if (!grouped[r.hub]) grouped[r.hub] = [];
        grouped[r.hub].push(v);
      });
      
      const list = Object.entries(grouped).map(([hub, vals]) => ({ hub, value: (vals.reduce((a, b) => a + b, 0) / vals.length) * 100 }));
      list.sort((a, b) => b.value - a.value);
      
      return ok({ top: list.slice(0, topN || 10), low: list.filter(item => item.value < CONFIG.TARGET_PCT).slice(-(lowN || 10)).reverse(), metricLabel: metric.label, segment: seg, mode: currentMode });
    } catch (e) { return fail(e.message); }
  });
}

function apiGetWeekly(mode, hubFilter, week, month) {
  const currentMode = mode || 'FM';
  const cacheKey = `weekly_v2_${currentMode}_${hubFilter || 'all'}_${week || '-'}_${month || '-'}`;
  return withApiCache(cacheKey, function() {
    try {
      const access = getUserAccess(Session.getActiveUser().getEmail(), currentMode);
      let finalHubFilter = hubFilter;
      if (access.hubAccess && !access.hubAccess.includes('Semua') && access.hubAccess.length > 0) {
        if (!finalHubFilter || !access.hubAccess.includes(finalHubFilter)) finalHubFilter = access.hubAccess[0];
      }

      let rows = filterByHub(currentMode === 'LM' ? getLMWeekly() : getFMWeekly(), finalHubFilter);
      if (week) rows = rows.filter(r => String(r.week) === String(week)); 
      else if (month && currentMode === 'FM') rows = rows.filter(r => r.month === month);
      
      rows.sort((a, b) => (a.score || 0) - (b.score || 0));

      const metricsList = currentMode === 'LM' ? CONFIG.LM_METRICS : CONFIG.FM_METRICS;
      let hasScores = false;
      const data = rows.map(r => {
        const obj = { Week: r.week, Month: r.month || '—', Hub: r.hub, Score: r.score !== null ? (r.score * 100).toFixed(2) + '%' : '—' };
        metricsList.forEach(m => {
          const v = r.metrics[m.key];
          obj[m.label] = v !== null ? (v * 100).toFixed(1) + '%' : '—';
        });

        /* Skor per metrik (poin, skala 0–100 agar jumlahnya sebanding dengan kolom Score).
           Sheet bisa menyimpan pecahan (0.1485) atau poin (14.85): skala ditentukan dengan membandingkan jumlahnya ke Score total. */
        const ms = r.metricScores || {};
        const vals = metricsList.map(m => ms[m.key]).filter(v => v !== null && v !== undefined && !isNaN(v));
        let factor = 100;
        if (vals.length && r.score !== null) {
          const sum = vals.reduce((a, b) => a + b, 0);
          if (Math.abs(sum - r.score * 100) < Math.abs(sum * 100 - r.score * 100)) factor = 1;
        }
        const sbm = {};
        metricsList.forEach(m => {
          const v = ms[m.key];
          if (v !== null && v !== undefined && !isNaN(v)) { sbm[m.label] = (v * factor).toFixed(2); hasScores = true; }
          else sbm[m.label] = '—';
        });
        obj.ScoreByMetric = sbm;
        return obj;
      });
      return ok({ headers: ['Week', 'Month', 'Hub', 'Score', ...metricsList.map(m => m.label)], rows: data, total: rows.length, mode: currentMode, hasScores: hasScores });
    } catch (e) { return fail(e.message); }
  });
}

function apiGetMonthly(mode, hubFilter, month) {
  const currentMode = mode || 'FM';
  if (currentMode === 'LM') return apiGetWeekly('LM', hubFilter, null, month);

  const cacheKey = `monthly_${currentMode}_${hubFilter || 'all'}_${month || '-'}`;
  return withApiCache(cacheKey, function() {
    try {
      const access = getUserAccess(Session.getActiveUser().getEmail(), currentMode);
      let finalHubFilter = hubFilter;
      if (access.hubAccess && !access.hubAccess.includes('Semua') && access.hubAccess.length > 0) {
        if (!finalHubFilter || !access.hubAccess.includes(finalHubFilter)) finalHubFilter = access.hubAccess[0];
      }

      let rows = filterByHub(getFMMonthly(), finalHubFilter);
      rows = filterByMonth(rows, month);
      rows.sort((a, b) => (a.score || 0) - (b.score || 0));

      const data = rows.map(r => {
        const obj = { Month: r.month, Hub: r.hub, Score: r.score !== null ? (r.score * 100).toFixed(2) + '%' : '—' };
        CONFIG.FM_METRICS.forEach(m => {
          const v = r.metrics[m.key];
          obj[m.label] = v !== null ? (v * 100).toFixed(1) + '%' : '—';
        });
        return obj;
      });
      return ok({ headers: ['Month', 'Hub', 'Score', ...CONFIG.FM_METRICS.map(m => m.label)], rows: data, total: rows.length, mode: currentMode });
    } catch (e) { return fail(e.message); }
  });
}

function apiGetPerHub(mode, segment, hubFilter, date, week, month) {
  const currentMode = mode || 'FM';
  const seg = segment || 'blended';
  const cacheKey = `perhub_${currentMode}_${seg}_${hubFilter || 'all'}_${date || '-'}_${week || '-'}_${month || '-'}`;
  return withApiCache(cacheKey, function() {
    try {
      const access = getUserAccess(Session.getActiveUser().getEmail(), currentMode);
      let finalHubFilter = hubFilter;
      if (access.hubAccess && !access.hubAccess.includes('Semua') && access.hubAccess.length > 0) {
        if (!finalHubFilter || !access.hubAccess.includes(finalHubFilter)) finalHubFilter = access.hubAccess[0];
      }

      let rows = currentMode === 'LM' ? getLMDaily() : (seg === 'high' ? getFMHigh() : (seg === 'regular' ? getFMRegular() : (getFMRegular() || []).concat(getFMHigh() || [])));
      rows = filterByHub(rows, finalHubFilter);
      rows = filterByDateWeekMonth(rows, date, week, month);
      
      const metricsList = currentMode === 'LM' ? CONFIG.LM_METRICS : CONFIG.FM_METRICS;
      const agg = {};
      rows.forEach(r => {
        metricsList.forEach(m => {
          const v = r.metrics[m.key];
          if (v === null || isNaN(v)) return;
          const k = r.hub + '||' + m.key;
          if (!agg[k]) agg[k] = { hub: r.hub, metric: m, vals: [] };
          agg[k].vals.push(v);
        });
      });
      
      const items = Object.values(agg).map(g => {
        const avg = g.vals.reduce((a, b) => a + b, 0) / g.vals.length;
        return {
          hub: g.hub, metricKey: g.metric.key, metricLabel: g.metric.label,
          achievement: avg * 100, target: 100, selisih: (avg - 1) * 100,
          status: achToStatus(avg), action: CONFIG.DEFAULT_ACTIONS[g.metric.key] || ''
        };
      });
      items.sort((a, b) => a.selisih - b.selisih);
      return ok({ items: items.slice(0, 500), segment: seg, total: items.length, mode: currentMode });
    } catch (e) { return fail(e.message); }
  });
}

function apiGetRekomendasiHub(mode, hubFilter, dateFilter, weekFilter, monthFilter, customThreshold) {
  const currentMode = mode || 'FM';
  // Menu Hypercare sementara hanya untuk City Lead / PIC
  const accessHc = getUserAccess(Session.getActiveUser().getEmail(), currentMode);
  if (!accessHc.isAllowed || !(accessHc.role === 'City Lead' || accessHc.role === 'PIC')) return fail('Menu Hypercare hanya untuk City Lead dan PIC.');
  const cutoff = getHypercareCutoffFor_(currentMode); // cutoff bersama, bukan per-browser
  const cacheKey = `rekomendasi_v19_${currentMode}_${dateFilter || '-'}_${weekFilter || '-'}_${monthFilter || '-'}_th_${cutoff}`;
  
  return withApiCache(cacheKey, function() {
    try {
      const metricsList = currentMode === 'LM' ? CONFIG.LM_METRICS : CONFIG.FM_METRICS;
      let rows = currentMode === 'LM' ? getLMWeekly() : getFMWeekly();
      
      if (weekFilter) rows = rows.filter(r => String(r.week) === String(weekFilter));
      else if (monthFilter && currentMode === 'FM') rows = rows.filter(r => String(r.month) === String(monthFilter));
      
      if (rows.length === 0) rows = (currentMode === 'LM' ? getLMWeekly() : getFMWeekly()).slice(-50);
      
      rows.sort((a, b) => (a.score || 0) - (b.score || 0));
      const top10 = rows.slice(0, 10);
      
      const savedInputs = getHypercareSavedInputs();
      const result = top10.map(r => {
        const issues = [];
        metricsList.forEach(m => {
          const v = r.metrics[m.key];
          const valPct = v !== null && v !== undefined && !isNaN(v) ? (isFraction_(Number(v)) ? Number(v) * 100 : Number(v)) : 0;
          
          if (valPct < cutoff) {
            const keyId = buildHypercareKey(r.hub, r.week, m.key);
            const saved = savedInputs[keyId] || {};

            issues.push({
              key: m.key, label: m.label, value: valPct, gap: cutoff - valPct,
              action: CONFIG.DEFAULT_ACTIONS[m.key] || 'Lakukan evaluasi pada proses ini.',
              rootCause: saved.rootCause || '', actionPlan: saved.actionPlan || '',
              updatedBy: saved.updatedBy || '', updatedAt: saved.updatedAt || '', status: saved.status || 'Open', dueDate: saved.dueDate || '', pic: saved.pic || ''
            });
          }
        });
        issues.sort((a, b) => b.gap - a.gap);
        return { hub: r.hub, group: r.week, score: r.score !== null ? (isFraction_(Number(r.score)) ? Number(r.score) * 100 : Number(r.score)) : null, issues };
      });
      
      return ok({ items: result, total: rows.length, source: 'weekly', activeWeek: weekFilter, activeMonth: monthFilter, mode: currentMode, targetCutoff: cutoff });
    } catch (e) { return fail(e.message); }
  });
}

function apiLogAccess() {
  try {
    const ss = getSS();
    let logSheet = ss.getSheetByName('Audit_Log');
    if (!logSheet) {
      logSheet = ss.insertSheet('Audit_Log');
      logSheet.appendRow(['Timestamp', 'Email', 'User Agent']);
      logSheet.getRange("A1:C1").setFontWeight("bold").setBackground("#f3f3f3");
      logSheet.setFrozenRows(1);
    }
    const email = Session.getActiveUser().getEmail() || 'Anonymous';
    const timestamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    logSheet.appendRow([timestamp, email, 'Web App Access']);
    return ok({ message: 'Access logged' });
  } catch (e) { return fail(e.message); }
}

function apiGetComprehensiveTrend(mode, viewType, metricKeyFilter, hubFilter) {
  const currentMode = mode || 'FM';
  const type = viewType || 'OVERALL';
  const cacheKey = `comp_trend_v18_${currentMode}_${type}_${metricKeyFilter || 'all'}_${hubFilter || 'all'}`;

  return withApiCache(cacheKey, function() {
    try {
      const userEmail = Session.getActiveUser().getEmail();
      const access = getUserAccess(userEmail, currentMode);

      if (type === 'OVERALL') {
        const isCLorPIC = access && (access.role === 'City Lead' || access.role === 'PIC');
        if (!isCLorPIC) {
          return fail('🔒 Akses Terbatas: Menu Trend Overall JTT-07 hanya dapat diakses oleh City Lead dan PIC.');
        }
      }

      let finalHubFilter = hubFilter;
      if (access.hubAccess && !access.hubAccess.includes('Semua') && access.hubAccess.length > 0) {
        if (!finalHubFilter || !access.hubAccess.includes(finalHubFilter)) {
          finalHubFilter = access.hubAccess[0];
        }
      }

      let rawWeekly = currentMode === 'LM' ? getLMWeekly() : getFMWeekly();
      if (!rawWeekly || !rawWeekly.length) return fail('Data mentah tidak ditemukan');

      rawWeekly = filterByHub(rawWeekly, finalHubFilter);

      function parseWeekToDate(weekStr) {
        if (!weekStr) return new Date(0);
        const monthMap = { jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11 };
        const match = String(weekStr).match(/(\d{1,2})\s*([a-zA-Z]{3})/i);
        if (match) {
          const day = parseInt(match[1], 10);
          const monthStr = match[2].toLowerCase();
          const month = monthMap[monthStr] !== undefined ? monthMap[monthStr] : 0;
          return inferYearDate_(month, day);
        }
        return new Date(0);
      }

      const weeksMap = {};
      rawWeekly.forEach(r => {
        if (r.week) {
          const wStr = String(r.week).trim();
          if (!weeksMap[wStr]) weeksMap[wStr] = parseWeekToDate(wStr);
        }
      });

      const sortedWeeks = Object.keys(weeksMap).sort((a, b) => weeksMap[a] - weeksMap[b]);
      const weeks = sortedWeeks.slice(-4);
      const weekSet = new Set(weeks);

      if (type === 'OVERALL') {
        const calcDivScore = (rows) => {
          const byWeek = {};
          rows.forEach(r => { const w = String(r.week).trim(); if (weekSet.has(w)) (byWeek[w] = byWeek[w] || []).push(r); });
          return weeks.map(w => byWeek[w] ? avgHubScore_(byWeek[w]).avg : null);
        };

        const divisionScores = calcDivScore(rawWeekly);
        const scopeLabel = currentMode === 'LM' ? 'Last-Mile (LM)' : 'First-Mile (FM)';

        const overallTrend = [
          { 
            scope: scopeLabel, 
            scores: divisionScores 
          }
        ];

        overallTrend.forEach(row => {
          const l0 = row.scores[weeks.length - 1];
          const l1 = row.scores[weeks.length - 2];
          row.delta = (l0 !== null && l1 !== null) ? (l0 - l1) : null;
        });

        return ok({ weeks, overallTrend, mode: currentMode });
      }

      if (type === 'ALL_HUB') {
        const hubAgg = {};
        rawWeekly.forEach(r => {
          const wStr = String(r.week).trim();
          const hub = r.hub || 'Unknown';
          if (!weekSet.has(wStr)) return;

          if (r.score !== undefined && r.score !== null && !isNaN(r.score)) {
            if (!hubAgg[hub]) hubAgg[hub] = {};
            if (!hubAgg[hub][wStr]) hubAgg[hub][wStr] = { sum: 0, count: 0 };
            hubAgg[hub][wStr].sum += Number(r.score);
            hubAgg[hub][wStr].count += 1;
          }
        });

        const trendData = Object.keys(hubAgg).map(hubName => {
          const weekScores = {};
          weeks.forEach((w, idx) => {
            const item = hubAgg[hubName][w];
            weekScores[`L${weeks.length - 1 - idx}W`] = (item && item.count > 0) ? (item.sum / item.count) * 100 : null;
          });

          const l0 = weekScores['L0W'];
          const l1 = weekScores['L1W'];
          const delta = (l0 !== null && l1 !== null) ? (l0 - l1) : null;

          return { hub: hubName, scores: weekScores, delta: delta };
        });

        trendData.sort((a, b) => {
          const dA = a.delta !== null && a.delta !== undefined ? a.delta : 999;
          const dB = b.delta !== null && b.delta !== undefined ? b.delta : 999;
          return dA - dB;
        });

        return ok({ weeks, trendData, mode: currentMode });
      }

      if (type === 'BY_METRIC') {
        const metricsList = currentMode === 'LM' ? CONFIG.LM_METRICS : CONFIG.FM_METRICS;
        const targetMetricKey = metricKeyFilter || metricsList[0].key;
        const targetMetric = metricsList.find(m => m.key === targetMetricKey) || metricsList[0];

        const hubAgg = {};
        rawWeekly.forEach(r => {
          const wStr = String(r.week).trim();
          const hub = r.hub || 'Unknown';
          if (!weekSet.has(wStr)) return;

          const v = r.metrics[targetMetric.key];
          if (v !== null && v !== undefined && !isNaN(v)) {
            if (!hubAgg[hub]) hubAgg[hub] = {};
            if (!hubAgg[hub][wStr]) hubAgg[hub][wStr] = { sum: 0, count: 0 };
            hubAgg[hub][wStr].sum += Number(v);
            hubAgg[hub][wStr].count += 1;
          }
        });

        const trendData = Object.keys(hubAgg).map(hubName => {
          const weekScores = {};
          weeks.forEach((w, idx) => {
            const item = hubAgg[hubName][w];
            weekScores[`L${weeks.length - 1 - idx}W`] = (item && item.count > 0) ? (item.sum / item.count) * 100 : null;
          });

          const l0 = weekScores['L0W'];
          const l1 = weekScores['L1W'];
          const delta = (l0 !== null && l1 !== null) ? (l0 - l1) : null;

          return { hub: hubName, metricLabel: targetMetric.label, scores: weekScores, delta: delta };
        });

        trendData.sort((a, b) => {
          const dA = a.delta !== null && a.delta !== undefined ? a.delta : 999;
          const dB = b.delta !== null && b.delta !== undefined ? b.delta : 999;
          return dA - dB;
        });

        return ok({ weeks, trendData, mode: currentMode, selectedMetric: targetMetric });
      }

    } catch (e) {
      return fail(e.message);
    }
  });
}

function apiGetWeeklyTrend(mode, hubFilter, lastNWeeks) {
  return apiGetComprehensiveTrend(mode, 'BY_METRIC', null, hubFilter);
}



function apiGetUserInfo() { try { return ok({ email: Session.getActiveUser().getEmail() || '' }); } catch (e) { return ok({ email: '' }); } }

/* Waktu data terakhir masuk = waktu terakhir file spreadsheet diubah (butuh scope Drive) */
function apiGetDataFreshness() {
  try {
    let t = null;
    try { t = DriveApp.getFileById(CONFIG.SPREADSHEET_ID).getLastUpdated(); } catch (e) {}
    return ok({ updatedAt: t ? Utilities.formatDate(t, CONFIG.TIMEZONE, 'dd MMM yyyy HH:mm') : '' });
  } catch (e) { return fail(e.message); }
}


/* Tahun dari tanggal saat ini (bukan 2026 permanen); label minggu >60 hari ke depan dianggap tahun lalu */
function inferYearDate_(m, d) {
  const now = new Date();
  let dt = new Date(now.getFullYear(), m, d);
  if (dt - now > 60 * 86400000) dt = new Date(now.getFullYear() - 1, m, d);
  return dt;
}

/* Cutoff Hypercare tersimpan di Script Properties → sama untuk semua pengguna */
function getHypercareCutoffFor_(mode) {
  const n = parseFloat(PropertiesService.getScriptProperties().getProperty('HC_CUTOFF_' + (mode === 'LM' ? 'LM' : 'FM')));
  return isNaN(n) ? (mode === 'LM' ? 98.00 : CONFIG.TARGET_PCT) : n;
}
function apiSetHypercareCutoff(mode, val) {
  try {
    const m = mode === 'LM' ? 'LM' : 'FM';
    const a = getUserAccess(Session.getActiveUser().getEmail(), m);
    if (!a.isAllowed || !(a.role === 'City Lead' || a.role === 'PIC')) return fail('Hanya City Lead/PIC yang dapat mengubah cutoff.');
    const n = parseFloat(val);
    if (isNaN(n) || n < 0 || n > 100) return fail('Nilai cutoff tidak valid.');
    PropertiesService.getScriptProperties().setProperty('HC_CUTOFF_' + m, String(n));
    return ok({ cutoff: n });
  } catch (e) { return fail(e.message); }
}

/* Simpan Hypercare: dikunci (anti baris ganda), divalidasi, dan dicek hak akses hub */
function apiSaveHypercareInput(hub, week, metricKey, rootCause, actionPlan, status, dueDate, pic, whyChain) {
  const mode = String(metricKey || '').indexOf('L_') === 0 ? 'LM' : 'FM';
  const access = getUserAccess(Session.getActiveUser().getEmail(), mode);
  if (!access.isAllowed) return fail('Akses ditolak.');
  if (access.hubAccess && !access.hubAccess.includes('Semua')) {
    const ok_ = access.hubAccess.some(h => cleanHubName(h) === cleanHubName(hub));
    if (!ok_) return fail('Anda tidak memiliki akses untuk menyimpan data hub ini.');
  }
  const allowed = ['Open', 'In progress', 'Menunggu validasi', 'Done'];
  status = allowed.indexOf(status) >= 0 ? status : 'Open';
  const gateMsg = hcvCheckStatus_(access, hub, week, metricKey, status);
  if (gateMsg) return fail(gateMsg);
  const clip = (s, n) => String(s || '').slice(0, n);
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return fail('Server sedang sibuk, coba lagi sebentar.'); }
  try {
    return saveHypercareInputUnsafe_(clip(hub, 100), clip(week, 50), clip(metricKey, 20), clip(rootCause, 500), clip(actionPlan, 2000), status, clip(dueDate, 10), clip(pic, 100),
      (whyChain === null || whyChain === undefined) ? null : clip(whyChain, 2500));
  } finally { lock.releaseLock(); }
}

/* Sheet Access Web di-cache 5 menit agar tidak dibaca ulang di setiap panggilan */
let ACCESS_ROWS_MEMO_ = null;
function getAccessRows_() {
  if (ACCESS_ROWS_MEMO_ === null) ACCESS_ROWS_MEMO_ = getAccessRowsRaw_();
  return ACCESS_ROWS_MEMO_;
}
function getAccessRowsRaw_() {
  const cache = CacheService.getScriptCache();
  const c = cache.get('ACCESS_ROWS');
  if (c) { try { return JSON.parse(c); } catch (e) {} }
  const sheet = getSS().getSheetByName(CONFIG.SHEETS.ACCESS_WEB);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  try { cache.put('ACCESS_ROWS', JSON.stringify(data), 300); } catch (e) {}
  return data;
}


/* Parameter tautan (?m=FM&p=weekly&hub=...) → state awal dashboard; hanya kunci yang diizinkan */
function sanitizeInitState_(e) {
  const p = (e && e.parameter) || {};
  if (p.m !== 'FM' && p.m !== 'LM') return '{}';
  const o = { m: p.m };
  ['p', 'hub', 'month', 'week', 'date', 'seg'].forEach(k => { if (p[k] !== undefined) o[k] = String(p[k]).slice(0, 100); });
  return JSON.stringify(o).replace(/</g, '\\u003c').replace(/>/g, '\\u003e').replace(/&/g, '\\u0026');
}
function apiGetAppUrl() { try { return ok({ url: ScriptApp.getService().getUrl() }); } catch (e) { return fail(e.message); } }

/* Selisih skor mingguan (minggu terpilih/terbaru vs minggu sebelumnya) */
function computeWow_(mode, hub, week) {
  try {
    const rows = filterByHub(mode === 'LM' ? getLMWeekly() : getFMWeekly(), hub);
    const mm = { jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11 };
    const parse = w => { const m = String(w).match(/(\d{1,2})\s*([a-zA-Z]{3})/); return m ? inferYearDate_(mm[m[2].toLowerCase()] || 0, parseInt(m[1], 10)) : new Date(0); };
    const byWeek = {};
    rows.forEach(r => { const w = String(r.week || '').trim(); if (w) (byWeek[w] = byWeek[w] || []).push(r); });
    const agg = {};
    Object.keys(byWeek).forEach(w => {
      const a = avgHubScore_(byWeek[w]);
      if (a.avg !== null) agg[w] = { v: a.avg, d: parse(w) };   // minggu template/kosong otomatis terbuang
    });
    const ws = Object.keys(agg).sort((a, b) => agg[a].d - agg[b].d);
    const i = week ? ws.indexOf(String(week).trim()) : ws.length - 1;
    if (i < 1) return null;
    return { week: ws[i], prevWeek: ws[i - 1], current: agg[ws[i]].v, prev: agg[ws[i - 1]].v, delta: agg[ws[i]].v - agg[ws[i - 1]].v };
  } catch (e) { return null; }
}

/* Minggu terbaru yang SUDAH ada datanya di sheet weekly (bukan minggu template kosong) */
function latestWeekLabel_(mode) {
  try {
    const mm = { jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11 };
    let best = null, bestT = -1;
    (mode === 'LM' ? getLMWeekly() : getFMWeekly()).forEach(r => {
      const w = String(r.week || '').trim();
      if (!w) return;
      const hasData = r.metrics && Object.keys(r.metrics).some(k => r.metrics[k] !== null && r.metrics[k] !== undefined);
      if (!hasData) return;
      const m = w.match(/(\d{1,2})\s*([a-zA-Z]{3})/);
      if (!m || mm[m[2].toLowerCase()] === undefined) return;
      const t = inferYearDate_(mm[m[2].toLowerCase()], parseInt(m[1], 10)).getTime();
      if (t > bestT) { bestT = t; best = w; }
    });
    return best;
  } catch (e) { return null; }
}

/* Snapshot landing page: skor mingguan terbaru per divisi (ikut hak akses hub pengguna) */
function apiGetLandingSnapshot() {
  return withApiCache('landing_snap_v3', function() {
    try {
      const email = Session.getActiveUser().getEmail();
      return ok({ FM: landingSnapshotFor_('FM', email), LM: landingSnapshotFor_('LM', email) });
    } catch (e) { return fail(e.message); }
  }, function(res) {
    // Jangan cache bila ada hub mencurigakan (skor tidak cocok dengan metriknya) atau salah satu divisi kosong
    const d = res && res.data; if (!d) return false;
    return ['FM', 'LM'].every(m => !d[m] || !d[m].hubsSuspect);
  });
}

function landingSnapshotFor_(mode, email) {
  try {
    const access = getUserAccess(email, mode);
    if (!access.isAllowed) return null;
    let rows = mode === 'LM' ? getLMWeekly() : getFMWeekly();
    if (access.hubAccess && access.hubAccess.length && !access.hubAccess.includes('Semua')) {
      const allow = access.hubAccess.map(cleanHubName);
      rows = rows.filter(r => allow.indexOf(cleanHubName(r.hub)) >= 0);
    }
    const mm = { jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11 };

    // Minggu dipakai hanya jika sudah terisi minimal sekian persen hub (minggu berjalan/template diabaikan)
    const MIN_COVERAGE = 0.5;

    const allHubs = {};
    const byWeek = {};
    rows.forEach(r => {
      const w = String(r.week || '').trim();
      const m = w.match(/(\d{1,2})\s*([a-zA-Z]{3})/);
      if (!m || mm[m[2].toLowerCase()] === undefined) return;
      allHubs[r.hub] = true;
      if (!byWeek[w]) byWeek[w] = { t: inferYearDate_(mm[m[2].toLowerCase()], parseInt(m[1], 10)).getTime(), hubs: {}, seen: {}, suspect: {} };
      byWeek[w].seen[r.hub] = true;
      const validScore = r.score !== null && r.score !== undefined && !isNaN(r.score);
      if (!validScore || !rowHasMetricData_(r)) return;
      if (rowScoreSuspect_(r)) { byWeek[w].suspect[r.hub] = true; return; }
      const v = scoreToPct_(r.score);
      (byWeek[w].hubs[r.hub] = byWeek[w].hubs[r.hub] || []).push(v);
    });

    const withData = Object.keys(byWeek).filter(w => Object.keys(byWeek[w].hubs).length > 0);
    if (!withData.length) return null;
    const minHubs = Math.ceil(Object.keys(allHubs).length * MIN_COVERAGE);
    let weeks = withData.filter(w => Object.keys(byWeek[w].hubs).length >= minHubs);
    if (!weeks.length) weeks = withData;               // cadangan: pakai minggu terbaru yang ada isinya
    weeks.sort((a, b) => byWeek[a].t - byWeek[b].t);

    const summarize = w => {
      const list = Object.keys(byWeek[w].hubs).map(h => {
        const a = byWeek[w].hubs[h];
        return { hub: h, score: a.reduce((x, y) => x + y, 0) / a.length };
      });
      return {
        list: list,
        avg: list.reduce((x, y) => x + y.score, 0) / list.length,
        noData: Object.keys(byWeek[w].seen).filter(h => !byWeek[w].hubs[h] && !byWeek[w].suspect[h]).length,
        suspect: Object.keys(byWeek[w].suspect).length
      };
    };
    const cur = summarize(weeks[weeks.length - 1]);
    const prev = weeks.length > 1 ? summarize(weeks[weeks.length - 2]) : null;
    cur.list.sort((a, b) => a.score - b.score);
    return {
      week: weeks[weeks.length - 1],
      score: cur.avg,
      prevScore: prev ? prev.avg : null,
      delta: prev ? cur.avg - prev.avg : null,
      hubCount: cur.list.length,
      hubsNoData: cur.noData,
      hubsSuspect: cur.suspect,
      hubsBelowWarn: cur.list.filter(x => x.score < CONFIG.WARN_PCT).length,
      worstHub: cur.list[0].hub,
      worstScore: cur.list[0].score
    };
  } catch (e) { return null; }
}
function debugScoreHubs(week) {
  week = week || '19 Oct - 25 Oct'; // samakan persis dengan label minggu di sheet
  const daily = new Set(getFMRegular().concat(getFMHigh()).map(r => r.hub));
  const byHub = {};
  getFMWeekly().filter(r => String(r.week) === String(week))
    .forEach(r => { (byHub[r.hub] = byHub[r.hub] || []).push(r); });
  const out = [];
  daily.forEach(h => {
    const rs = byHub[h] || [];
    let why;
    if (!rs.length) why = 'TIDAK ADA baris di FM Blended Weekly';
    else if (rs.some(rowScoreUsable_)) why = 'dipakai';
    else if (rs.every(r => r.score === null || isNaN(r.score))) why = 'SCORE kosong/error';
    else if (!rs.some(rowHasMetricData_)) why = 'semua metrik kosong/0';
    else why = 'SCORE dianggap janggal';
    out.push(h + ' | ' + why + ' | ' + rs.map(r => r.score).join(','));
  });
  Logger.log(out.sort().join('\n'));
  Logger.log('Ada di weekly tapi tidak di daily: ' + Object.keys(byHub).filter(h => !daily.has(h)).join(', '));
}
function diagHubScore() {
  const weekLabel = '19 Oct - 25 Oct'; // sesuaikan
  const all = getHubList('FM');
  const rows = getFMWeekly().filter(r => String(r.week) === weekLabel);
  const usable = new Set(rows.filter(rowScoreUsable_).map(r => r.hub));
  const inSheet = new Set(rows.map(r => r.hub));
  Logger.log('Total hub (daftar): ' + all.length + ' | di sheet mingguan: ' + inSheet.size + ' | terpakai: ' + usable.size);
  Logger.log('Tidak ada baris mingguan: ' + all.filter(h => !inSheet.has(h)).join(', '));
  rows.filter(r => !rowScoreUsable_(r)).forEach(r =>
    Logger.log('Dibuang: ' + r.hub + ' | score=' + r.score +
      ' | punyaMetrik=' + rowHasMetricData_(r) + ' | janggal=' + rowScoreSuspect_(r)));
}