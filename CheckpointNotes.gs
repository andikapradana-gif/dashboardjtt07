/************************************************************
 * CHECKPOINTNOTES.GS — Menu Checkpoint: catatan kendala/need support
 * per hub & metrik + angka progress untuk notulen. Hanya CL / PIC.
 * Acuan = rencana Hypercare terakhir di sheet Hypercare_Plan.
 *
 * PERFORMA (v2):
 *  - Angka progress dibaca dari BLOK KOLOM mingguan saja (bukan seluruh sheet)
 *    dan di-cache 20 menit. "Muat ulang data" di dashboard ikut membersihkannya.
 *  - Catatan tersimpan memperbarui cache di tempat (tidak dihapus), jadi memuat
 *    papan setelah mengetik tidak membaca sheet lagi.
 *  - Simpan catatan hanya membaca kolom A untuk mencari baris.
 *  - Jalankan installWarmTrigger() SEKALI agar cache selalu hangat (tiap 10 menit).
 ************************************************************/

const CKP_NOTES_SHEET = 'Checkpoint_Notes';
const CKP_LOG_SHEET = 'Checkpoint_Notulen';
const CKP_CORE_TTL = 1200;    // detik: cache angka baseline vs terbaru
const CKP_NOTES_TTL = 1200;   // detik: cache catatan per rencana

/* Rencana tidak pernah berubah setelah dicatat → cache 6 jam per planId (menyimpan catatan jadi tanpa baca sheet) */
function ckpFindPlan_(planId) {
  const cache = CacheService.getScriptCache(), ck = 'CKP_PLAN_' + planId;
  const c = cache.get(ck);
  if (c) { try { return JSON.parse(c); } catch (e) {} }
  const p = ckpFindPlanUncached_(planId);
  if (p) { try { cache.put(ck, JSON.stringify(p), 21600); } catch (e) {} }
  return p;
}

function ckpFindPlanUncached_(planId) {
  const modes = ['FM', 'LM'];
  for (let k = 0; k < modes.length; k++) {   // rencana terbaru sudah ada di cache → tanpa baca sheet
    const p = cpLatestPlan_(modes[k]);
    if (p && p.planId === String(planId)) return { planId: p.planId, mode: modes[k], week: p.week, focus: p.focus || [] };
  }
  const sh = getSS().getSheetByName(HC_PLAN_SHEET);   // rencana lama
  if (!sh || sh.getLastRow() < 2) return null;
  const data = sh.getRange(1, 1, sh.getLastRow(), 15).getValues();   // kolom A–O saja (tanpa Message)
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(planId) && String(data[i][14] || 'Hypercare') === 'Hypercare') {
      let focus = [];
      try { focus = JSON.parse(String(data[i][6] || '[]')); } catch (e) {}
      return { planId: String(data[i][0]), mode: String(data[i][3]) === 'LM' ? 'LM' : 'FM', week: String(data[i][4] || ''), focus: focus };
    }
  }
  return null;
}

function ckpNotesMap_(planId) {
  const cache = CacheService.getScriptCache(), ck = 'CKP_NOTES_' + planId;
  const c = cache.get(ck);
  if (c) { try { return JSON.parse(c); } catch (e) {} }
  const map = {};
  const sh = getSS().getSheetByName(CKP_NOTES_SHEET);
  if (sh && sh.getLastRow() >= 2) {
    sh.getDataRange().getValues().slice(1).forEach(r => {
      if (String(r[1]) !== String(planId)) return;
      map[cleanHubName(r[3]) + '|' + String(r[4])] = {
        kendala: r[5] || '', support: r[6] || '', pic: r[7] || '', status: r[8] || 'Open',
        updatedBy: r[9] || '', updatedAt: r[10] ? (r[10] instanceof Date ? Utilities.formatDate(r[10], CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm') : String(r[10])) : ''
      };
    });
  }
  try { cache.put(ck, JSON.stringify(map), CKP_NOTES_TTL); } catch (e) {}
  return map;
}

/* Baca HANYA blok kolom mingguan (WEEK…SCORE…metrik), bukan seluruh sheet.
   Baris yang dihasilkan sama dengan getFMWeekly()/getLMWeekly() (memakai normalizer yang sama). */
function ckpWeeklyRows_(m) {
  const isLM = m === 'LM';
  const c = isLM ? CONFIG.LM_COLUMNS.WEEKLY : CONFIG.FM_COLUMNS.WEEKLY;
  const idx = isLM
    ? [c.WEEK, c.HUB, c.SCORE].concat(c.METRIC_INDICES)
    : [c.WEEK, c.MONTH, c.HUB, c.SCORE].concat(Array.from({ length: c.METRICS_COUNT }, (_, i) => c.METRICS_START + i));
  const c0 = Math.min.apply(null, idx), c1 = Math.max.apply(null, idx);
  const name = isLM ? CONFIG.SHEETS.LM_BACKUP : CONFIG.SHEETS.FM_WEEKLY;
  const sh = getSS().getSheetByName(name);
  if (!sh) throw new Error('Sheet tidak ditemukan: ' + name);
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < c0 + 1) return { rows: [], transient: false };
  const width = Math.min(c1, lastCol - 1) - c0 + 1;
  // Sheet LM juga memuat data harian yang jauh lebih panjang: cari baris terakhir yang berisi MINGGU, lalu baca blok sampai di situ saja
  const wk = sh.getRange(2, c.WEEK + 1, lastRow - 1, 1).getValues();
  let lastWeekly = 0;
  for (let i = wk.length - 1; i >= 0; i--) { if (wk[i][0] !== '' && wk[i][0] !== null) { lastWeekly = i + 1; break; } }
  if (!lastWeekly) return { rows: [], transient: false };
  const block = sh.getRange(2, c0 + 1, lastWeekly, width).getValues();
  const wi = c.WEEK - c0, hi = c.HUB - c0;
  const isTransient = v => typeof v === 'string' && (v === 'Loading...' || v === 'Loading…' || /^#(N\/A|REF!|VALUE!|ERROR!|DIV\/0!|NAME\?)/.test(v));
  const rows = []; let transient = false;
  for (let i = 0; i < block.length; i++) {
    const b = block[i];
    if (!b[wi] || !b[hi]) continue;                       // baris kosong di blok mingguan
    if (!transient && b.some(isTransient)) transient = true;
    const raw = [];
    for (let k = 0; k < width; k++) raw[c0 + k] = b[k];   // indeks mutlak seperti di CONFIG
    const r = isLM ? normalizeLMWeeklyRow(raw) : normalizeFMWeeklyRow(raw);
    if (r.hub && r.week) rows.push(r);
  }
  return { rows: rows, transient: transient };
}

/* Hitung angka baseline vs terbaru untuk satu rencana (tanpa cek akses, tanpa cache) */
function ckpComputeCore_(m, plan, weekFilter, cutoff) {
  const tc = Date.now();
  const metricsList = m === 'LM' ? CONFIG.LM_METRICS : CONFIG.FM_METRICS;
  const label = {}; metricsList.forEach(x => { label[x.key] = x.label; });
  const tr = Date.now();
  const rd = ckpWeeklyRows_(m);
  const readMs = Date.now() - tr;
  const all = rd.rows.filter(r => r.hub && r.week && rowHasMetricData_(r));

  const weekSet = {}; all.forEach(r => { weekSet[String(r.week).trim()] = true; });
  const weeks = Object.keys(weekSet).sort((a, b) => hcv2WeekDate_(a) - hcv2WeekDate_(b));
  if (!weeks.length) return { res: fail('Belum ada data mingguan yang terisi.'), transient: rd.transient };

  const baseWeek = String(plan.week).trim();
  let curWeek = String(weekFilter || '').trim();
  if (!curWeek || weeks.indexOf(curWeek) < 0) curWeek = weeks[weeks.length - 1];

  const base = {}, cur = {};   // satu putaran data untuk dua minggu (baris terakhir menang)
  all.forEach(r => {
    const w = String(r.week).trim();
    if (w === baseWeek) base[cleanHubName(r.hub)] = r;
    if (w === curWeek) cur[cleanHubName(r.hub)] = r;
  });
  const pctOf = v => (v === null || v === undefined || isNaN(v)) ? null : scoreToPct_(v);

  const items = plan.focus.map(f => {
    const ch = cleanHubName(f.hub), b = base[ch], c = cur[ch];
    const scoreB = b ? pctOf(b.score) : null, scoreC = c ? pctOf(c.score) : null;
    const mets = (f.metrics || []).map(k => {
      const vb = b ? pctOf(b.metrics[k]) : null, vc = c ? pctOf(c.metrics[k]) : null;
      return { key: k, label: label[k] || k, base: vb, cur: vc, delta: (vb !== null && vc !== null) ? vc - vb : null, reached: vc !== null && vc >= cutoff };
    });
    return { hub: f.hub, scoreBase: scoreB, scoreCur: scoreC, scoreDelta: (scoreB !== null && scoreC !== null) ? scoreC - scoreB : null, metrics: mets };
  });
  const res = ok({ mode: m, planId: plan.planId, planTitle: plan.title, planDate: plan.date, baseWeek: baseWeek, curWeek: curWeek, cutoff: cutoff, sameWeek: baseWeek === curWeek, items: items, computeMs: Date.now() - tc, readMs: readMs });
  return { res: res, transient: rd.transient, weeks: weeks };
}

/* Cache sendiri (20 menit). Kunci memuat API_VER → tombol "Muat ulang data" otomatis membersihkannya.
   Tanpa accessScope_(): menu ini hanya bisa dibuka CL/PIC dan angkanya sama untuk semuanya.
   force = true → lewati cache dan hitung ulang (dipakai trigger pemanasan). */
function ckpCoreCached_(m, plan, weekFilter, cutoff, force) {
  const cache = CacheService.getScriptCache();
  const ver = cache.get('API_VER') || '0';
  const keyFor = w => ['CKP_CORE_v2', ver, m, plan.planId, w || '-', cutoff].join('|');
  const wkKey = ['CKP_WEEKS', ver, m].join('|');
  let wf = String(weekFilter || '').trim();

  // Minggu dari filter dashboard bisa tidak ada di sheet mingguan (mis. minggu berjalan dari data harian).
  // Hasilnya sama dengan "minggu terbaru", jadi pakai kunci yang sama agar kena cache yang dipanaskan trigger.
  let known = null;
  try { known = JSON.parse(cache.get(wkKey) || 'null'); } catch (e) {}
  if (known && wf && known.indexOf(wf) < 0) wf = '';

  if (!force) {
    const hit = cacheGetBig_(keyFor(wf));
    if (hit) { try { return { res: JSON.parse(hit), computed: false }; } catch (e) {} }
  }

  let out;
  try { out = ckpComputeCore_(m, plan, wf, cutoff); } catch (e) { return { res: fail(e.message), computed: true }; }
  const res = out.res;
  if (res && res.success && !out.transient) {   // sheet yang sedang menghitung ulang tidak disimpan
    try {
      if (out.weeks) cache.put(wkKey, JSON.stringify(out.weeks), CKP_CORE_TTL);
      const s = JSON.stringify(res);
      const k0 = (wf && out.weeks && out.weeks.indexOf(wf) < 0) ? '' : wf;
      cachePutBig_(keyFor(k0), s, CKP_CORE_TTL);
      const cw = res.data.curWeek;
      if (cw && cw !== k0) cachePutBig_(keyFor(cw), s, CKP_CORE_TTL);   // "minggu terbaru" dan "minggu itu" berbagi hasil
    } catch (e) {}
  }
  return { res: res, computed: true };
}

/* Papan checkpoint: angka baseline vs terbaru (di-cache) + catatan tersimpan (selalu terbaru) */
function apiGetCheckpointBoard(mode, weekFilter) {
  const m = mode === 'LM' ? 'LM' : 'FM';
  const T0 = Date.now(), tm = {};
  if (!cpRequireLead_(m)) return fail('Menu Checkpoint hanya untuk City Lead dan PIC.');
  tm.akses = Date.now() - T0;
  try {
    const plan = cpLatestPlan_(m);
    tm.rencana = Date.now() - T0;
    if (!plan) return fail('Belum ada meeting Hypercare ' + m + '. Buat dulu lewat Hypercare v2 → "Blast & Google Meet".');
    if (!plan.focus || !plan.focus.length) return fail('Rencana Hypercare terakhir tidak memuat daftar metrik fokus.');
    const cutoff = getHypercareCutoffFor_(m);

    const cr = ckpCoreCached_(m, plan, weekFilter, cutoff, false);
    const core = cr.res, computed = cr.computed;
    tm.inti = Date.now() - T0;
    if (!core || !core.success) return core;

    // gabungkan catatan + hitung ringkasan (murah)
    const d = core.data, notes = ckpNotesMap_(plan.planId);
    let total = 0, reached = 0, improved = 0, worsened = 0; const sb = [], sc = [];
    d.items.forEach(h => {
      const ch = cleanHubName(h.hub);
      if (h.scoreBase !== null) sb.push(h.scoreBase); if (h.scoreCur !== null) sc.push(h.scoreCur);
      h.metrics.forEach(mt => {
        const n = notes[ch + '|' + mt.key] || {};
        mt.kendala = n.kendala || ''; mt.support = n.support || ''; mt.pic = n.pic || ''; mt.status = n.status || 'Open';
        mt.updatedBy = n.updatedBy || ''; mt.updatedAt = n.updatedAt || '';
        total++; if (mt.reached) reached++;
        if (mt.delta !== null && mt.delta > 0.005) improved++; if (mt.delta !== null && mt.delta < -0.005) worsened++;
      });
    });
    const avg = a => a.length ? a.reduce((x, y) => x + y, 0) / a.length : null;
    const avgB = avg(sb), avgC = avg(sc);
    d.summary = { total: total, reached: reached, improved: improved, worsened: worsened, avgBase: avgB, avgCur: avgC, avgDelta: (avgB !== null && avgC !== null) ? avgC - avgB : null };
    tm.catatan = Date.now() - T0;
    d.timing = { serverMs: Date.now() - T0, steps: tm, fromCache: !computed, computeMs: d.computeMs || 0, readMs: d.readMs || 0 };
    return core;
  } catch (e) { return fail(e.message); }
}

/* Dipanggil trigger waktu: menghitung ulang & menyimpan cache sebelum kedaluwarsa (tanpa sesi pengguna) */
function warmCheckpointCache() {
  ['FM', 'LM'].forEach(m => {
    try {
      const plan = cpLatestPlan_(m);
      if (!plan || !plan.focus || !plan.focus.length) return;
      ckpCoreCached_(m, plan, '', getHypercareCutoffFor_(m), true);
      ckpNotesMap_(plan.planId);
    } catch (e) { Logger.log('warmCheckpointCache ' + m + ': ' + e.message); }
  });
}

/* Jalankan SEKALI secara manual dari editor Apps Script untuk memasang trigger 10 menit */
function installWarmTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'warmCheckpointCache')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('warmCheckpointCache').timeBased().everyMinutes(10).create();
}

/* Simpan / perbarui catatan satu hub + metrik */
function apiSaveCheckpointNote(planId, hub, metricKey, kendala, support, pic, status) {
  const plan = ckpFindPlan_(planId);
  if (!plan) return fail('Rencana Hypercare tidak ditemukan.');
  if (!cpRequireLead_(plan.mode)) return fail('Hanya City Lead/PIC.');
  const ch = cleanHubName(hub);
  const f = plan.focus.filter(x => cleanHubName(x.hub) === ch)[0];
  if (!f || (f.metrics || []).indexOf(String(metricKey)) < 0) return fail('Hub/metrik tidak termasuk rencana Hypercare ini.');
  const allowed = ['Open', 'In progress', 'Done'];
  status = allowed.indexOf(status) >= 0 ? status : 'Open';
  const clip = (s, n) => String(s || '').slice(0, n);

  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return fail('Server sedang sibuk, coba lagi sebentar.'); }
  try {
    const ss = getSS();
    let sh = ss.getSheetByName(CKP_NOTES_SHEET);
    if (!sh) {
      sh = ss.insertSheet(CKP_NOTES_SHEET);
      sh.appendRow(['Key_ID', 'Plan_ID', 'Mode', 'Hub', 'Metric_Key', 'Kendala', 'Need_Support', 'PIC', 'Status', 'Updated_By', 'Updated_At']);
      sh.getRange('A1:K1').setFontWeight('bold').setBackground('#f3f3f3'); sh.setFrozenRows(1);
    }
    const email = Session.getActiveUser().getEmail() || 'Anonymous';
    const nowStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    const keyId = planId + '|' + ch + '|' + metricKey;
    const row = [keyId, planId, plan.mode, f.hub, metricKey, hcSafeCell_(clip(kendala, 1500)), hcSafeCell_(clip(support, 1500)), hcSafeCell_(clip(pic, 100)), status, email, nowStr];

    // cari baris hanya lewat kolom A (bukan seluruh sheet)
    const lastRow = sh.getLastRow();
    let idx = -1;
    if (lastRow >= 2) {
      const keys = sh.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < keys.length; i++) { if (String(keys[i][0]) === keyId) { idx = i + 2; break; } }
    }
    if (idx > 0) sh.getRange(idx, 1, 1, row.length).setValues([row]); else sh.appendRow(row);

    // perbarui cache di tempat agar pemuatan berikutnya tidak membaca sheet lagi
    const cache = CacheService.getScriptCache();
    try {
      const map = ckpNotesMap_(planId);
      map[ch + '|' + String(metricKey)] = { kendala: clip(kendala, 1500), support: clip(support, 1500), pic: clip(pic, 100), status: status, updatedBy: email, updatedAt: nowStr.slice(0, 16) };
      cache.put('CKP_NOTES_' + planId, JSON.stringify(map), CKP_NOTES_TTL);
    } catch (e) { try { cache.remove('CKP_NOTES_' + planId); } catch (e2) {} }
    return ok({ saved: true });
  } catch (e) { return fail(e.message); } finally { lock.releaseLock(); }
}

/* Arsipkan notulen yang sudah jadi */
function apiSaveNotulen(planId, progressWeek, text) {
  const plan = ckpFindPlan_(planId);
  if (!plan) return fail('Rencana Hypercare tidak ditemukan.');
  if (!cpRequireLead_(plan.mode)) return fail('Hanya City Lead/PIC.');
  const body = String(text || '').trim();
  if (!body) return fail('Notulen kosong.');
  try {
    const ss = getSS();
    let sh = ss.getSheetByName(CKP_LOG_SHEET);
    if (!sh) {
      sh = ss.insertSheet(CKP_LOG_SHEET);
      sh.appendRow(['Created_At', 'Created_By', 'Plan_ID', 'Mode', 'Hypercare_Week', 'Progress_Week', 'Notulen']);
      sh.getRange('A1:G1').setFontWeight('bold').setBackground('#f3f3f3'); sh.setFrozenRows(1);
    }
    sh.appendRow([Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'), Session.getActiveUser().getEmail() || 'Anonymous', planId, plan.mode,
      hcSafeCell_(plan.week), hcSafeCell_(String(progressWeek || '').slice(0, 50)), hcSafeCell_(body.slice(0, 20000))]);
    return ok({ saved: true });
  } catch (e) { return fail(e.message); }
}