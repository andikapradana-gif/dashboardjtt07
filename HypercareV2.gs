/************************************************************
 * HYPERCAREV2.GS — Penentuan otomatis hub & metrik Hypercare
 * Menu: City Lead & PIC saja.
 *
 * Logika:
 *  - Per hub & metrik di minggu terpilih: gap = cutoff - nilai (hanya yang di bawah cutoff).
 *  - Dampak ke skor hub (poin)  = bobot metrik × gap ÷ total bobot.
 *  - Per hub dipilih maksimal N metrik dengan dampak terbesar (fokus).
 *  - Hub diurutkan dari potensi kenaikan terbesar → top N hub = daftar Hypercare.
 *  - Potensi kenaikan KPI JTT-07 = Σ potensi hub ÷ jumlah hub (estimasi).
 ************************************************************/

function hcv2WeekDate_(w) {
  const mm = { jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11 };
  const m = String(w || '').match(/(\d{1,2})\s*([a-zA-Z]{3})/);
  return (m && mm[m[2].toLowerCase()] !== undefined) ? inferYearDate_(mm[m[2].toLowerCase()], parseInt(m[1], 10)) : new Date(0);
}

function hcv2Clamp_(v, lo, hi, def) {
  const n = parseInt(v, 10);
  return isNaN(n) ? def : Math.max(lo, Math.min(hi, n));
}

/* Bobot metrik dari sheet "Matriks & Bobot", dicocokkan berurutan:
   key → label persis → label ternormalisasi → kode "(A1)" di label → kode dari key.
   Bobot 0 sah (metrik tidak berpengaruh). Tidak ketemu → 1 dan ditandai matched:false. */
function hcv2Norm_(s) { return String(s || '').toLowerCase().replace(/[^a-z0-9]/g, ''); }
function hcv2Bare_(s) { return hcv2Norm_(String(s || '').replace(/^\s*\(.*?\)\s*/, '')); }
function hcv2Code_(s) { const m = String(s || '').match(/\(\s*([A-Za-z]+\d+)\s*\)/); return m ? m[1].toLowerCase() : ''; }
function hcv2ResolveWeights_(mode, metricsList) {
  const raw = (getMetricWeights()[mode]) || {};
  const byNorm = {}, byCode = {}, byBare = {};
  Object.keys(raw).forEach(k => {
    const n = hcv2Norm_(k); if (n && byNorm[n] === undefined) byNorm[n] = raw[k];
    const b = hcv2Bare_(k); if (b && byBare[b] === undefined) byBare[b] = raw[k];
    const c = hcv2Code_(k) || (/^[A-Za-z]+\d+$/.test(String(k).trim()) ? String(k).trim().toLowerCase() : '');
    if (c && byCode[c] === undefined) byCode[c] = raw[k];
  });
  const map = {}, info = [];
  metricsList.forEach(mt => {
    const keyCode = String(mt.key).replace(/^L_/i, '').toLowerCase();
    const tries = [
      ['key', raw[mt.key]], ['label', raw[mt.label]], ['nama', byNorm[hcv2Norm_(mt.label)]],
      ['nama tanpa kode', byBare[hcv2Bare_(mt.label)]],
      ['kode label', byCode[hcv2Code_(mt.label)]], ['kode key', byCode[keyCode]]
    ];
    let w = 1, src = 'bawaan (tidak ditemukan)', matched = false;
    for (let i = 0; i < tries.length; i++) {
      const v = tries[i][1];
      if (v !== undefined && v !== null && isFinite(Number(v)) && Number(v) >= 0) { w = Number(v); src = tries[i][0]; matched = true; break; }
    }
    map[mt.key] = w;
    info.push({ key: mt.key, label: mt.label, weight: w, source: src, matched: matched });
  });
  if (Object.keys(map).every(k => map[k] === 0)) {   // semua 0 → tak bisa dibandingkan
    info.forEach(i => { i.weight = 1; i.source = 'bawaan (semua bobot 0)'; i.matched = false; map[i.key] = 1; });
  }
  return { map: map, info: info };
}

function apiGetHypercareV2(mode, weekFilter, topN, maxMetrics) {
  const m = mode === 'LM' ? 'LM' : 'FM';
  const access = getUserAccess(Session.getActiveUser().getEmail(), m);
  if (!access.isAllowed) return fail('Akses ditolak.');
  // Logika akses seperti Hypercare lama: CL/PIC (atau akses "Semua") melihat seluruh daftar,
  // Hub Lead hanya melihat hubnya sendiri bila hub tersebut masuk daftar prioritas.
  const isFull = access.role === 'City Lead' || access.role === 'PIC' || (access.hubAccess || []).indexOf('Semua') >= 0;
  const n = hcv2Clamp_(topN, 3, 20, 10);
  const mx = hcv2Clamp_(maxMetrics, 1, 5, 3);
  const cutoff = getHypercareCutoffFor_(m);
  const cacheKey = `hcv2b_${m}_${weekFilter || '-'}_${n}_${mx}_th_${cutoff}`;

  let res = withApiCache(cacheKey, function () {
    try {
      const metricsList = m === 'LM' ? CONFIG.LM_METRICS : CONFIG.FM_METRICS;
      const all = (m === 'LM' ? getLMWeekly() : getFMWeekly()).filter(r => r.hub && r.week && rowHasMetricData_(r));
      if (!all.length) return fail('Belum ada data mingguan yang terisi.');

      // urutan minggu yang sudah punya data
      const weekSet = {};
      all.forEach(r => { weekSet[String(r.week).trim()] = true; });
      const weeks = Object.keys(weekSet).sort((a, b) => hcv2WeekDate_(a) - hcv2WeekDate_(b));
      let week = String(weekFilter || '').trim(), fallback = false;
      if (!week || weeks.indexOf(week) < 0) { fallback = !!week; week = weeks[weeks.length - 1]; }
      const idx = weeks.indexOf(week);
      const prevWeek = idx > 0 ? weeks[idx - 1] : '';

      const byHub = rows => { const o = {}; rows.forEach(r => { o[cleanHubName(r.hub)] = r; }); return o; };
      const cur = byHub(all.filter(r => String(r.week).trim() === week));
      const prev = prevWeek ? byHub(all.filter(r => String(r.week).trim() === prevWeek)) : {};

      const wInfo = hcv2ResolveWeights_(m, metricsList);
      const wOf = mt => wInfo.map[mt.key];
      const pctOf = v => (v === null || v === undefined || isNaN(v)) ? null : scoreToPct_(v);

      const hubs = [];
      Object.keys(cur).forEach(k => {
        const r = cur[k], p = prev[k];
        const issues = [];
        // total bobot hanya dari metrik yang punya nilai di hub ini
        let availW = 0;
        metricsList.forEach(mt => { if (pctOf(r.metrics[mt.key]) !== null) availW += wOf(mt); });
        if (!availW) return;
        metricsList.forEach(mt => {
          const val = pctOf(r.metrics[mt.key]);
          if (val === null || val >= cutoff) return;
          const pv = p ? pctOf(p.metrics[mt.key]) : null;
          const gap = cutoff - val;
          issues.push({
            key: mt.key, label: mt.label, value: val, gap: gap,
            weight: wOf(mt), weightPct: wOf(mt) / availW * 100, impact: wOf(mt) * gap / availW,
            prevValue: pv, delta: pv === null ? null : val - pv,
            action: CONFIG.DEFAULT_ACTIONS[mt.key] || 'Lakukan evaluasi pada proses ini.'
          });
        });
        if (!issues.length) return;
        issues.sort((a, b) => b.impact - a.impact);
        const focus = issues.slice(0, mx);
        hubs.push({
          hub: r.hub, group: week,
          score: (r.score === null || r.score === undefined || isNaN(r.score)) ? null : scoreToPct_(r.score),
          belowCount: issues.length,
          gain: focus.reduce((a, i) => a + i.impact, 0),
          gainAll: issues.reduce((a, i) => a + i.impact, 0),
          issues: focus
        });
      });

      hubs.sort((a, b) => b.gain - a.gain || (a.score || 0) - (b.score || 0));
      const picked = hubs.slice(0, n);
      const hubCount = Object.keys(cur).length;
      const totalGain = picked.reduce((a, h) => a + h.gain, 0);
      const currentAvg = avgHubScore_(Object.keys(cur).map(k => cur[k])).avg;
      const projectedAvg = currentAvg === null ? null : Math.min(100, currentAvg + totalGain / (hubCount || 1));

      return ok({
        mode: m, week: week, prevWeek: prevWeek, weekFallback: fallback,
        cutoff: cutoff, topN: n, maxMetrics: mx,
        hubCount: hubCount, hubsBelowCutoff: hubs.length,
        currentAvg: currentAvg, projectedAvg: projectedAvg,
        totalGain: totalGain / (hubCount || 1),
        weights: wInfo.info,
        items: picked
      });
    } catch (e) { return fail(e.message); }
  });

  // data tersimpan digabung SETELAH cache, agar root cause / action plan selalu terbaru
  if (res && res.success && res.data && res.data.items) {
    if (!isFull) {
      const mine = (access.hubAccess || []).map(cleanHubName);
      const d0 = res.data;
      res = { success: true, error: null, meta: res.meta || {}, data: Object.assign({}, d0, {
        restricted: true,
        items: d0.items.filter(h => mine.indexOf(cleanHubName(h.hub)) >= 0),
        weights: [], currentAvg: null, projectedAvg: null, totalGain: 0, hubsBelowCutoff: 0, hubCount: 0,
        userHubs: access.hubAccess || []
      }) };
    }
    const saved = getHypercareSavedInputs();
    res.data.items.forEach(h => h.issues.forEach(i => {
      const s = saved[buildHypercareKey(h.hub, h.group, i.key)] || {};
      i.rootCause = s.rootCause || ''; i.whyChain = s.whyChain || ''; i.actionPlan = s.actionPlan || '';
      i.status = s.status || 'Open'; i.dueDate = s.dueDate || ''; i.pic = s.pic || '';
      i.updatedBy = s.updatedBy || ''; i.updatedAt = s.updatedAt || '';
      i.temuan = s.temuan || {}; i.val = s.val || {};
    }));
  }
  return res;
}