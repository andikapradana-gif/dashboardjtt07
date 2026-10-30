/************************************************************
 * HYPERCAREBLAST.GS — Guest, Google Meet & log rencana Hypercare
 * Hanya City Lead / PIC.
 * PERLU: Services → tambahkan "Google Calendar API" (Advanced Service, id: Calendar)
 *        dan scope https://www.googleapis.com/auth/calendar di appsscript.json.
 ************************************************************/

const HC_PLAN_SHEET = 'Hypercare_Plan';
const HC_TZ = 'Asia/Jakarta';

/* Undangan Meet hanya ke Hub Lead: kolom Role (C) di sheet Access Web harus mengandung "Hub Lead" */
const HC_GUEST_ROLE_RE = /hub\s*lead/i;

function hcIsEmail_(s) { return /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(String(s || '')); }
function hcSafeCell_(v) { const s = String(v == null ? '' : v); return /^[=+\-@]/.test(s) ? "'" + s : s; }

/* Email Hub Lead (dari sheet Access Web) untuk daftar hub tertentu */
function hcGuestsForHubs_(mode, hubs) {
  const data = getAccessRows_() || [];
  const targets = hubs.map(cleanHubName);
  const guests = {}, covered = {};
  for (let i = 1; i < data.length; i++) {
    const email = String(data[i][1] || '').toLowerCase().trim();
    if (!hcIsEmail_(email)) continue;
    if (!HC_GUEST_ROLE_RE.test(String(data[i][2] || ''))) continue;   // lewati City Lead, PIC, dan role lain
    const list = mode === 'LM' ? [data[i][4], data[i][5]] : String(data[i][3] || '').split(',');
    list.map(s => String(s || '').trim()).filter(s => s && s !== '-' && s.toLowerCase() !== 'null').forEach(h => {
      const c = cleanHubName(h);
      if (targets.indexOf(c) >= 0) { (guests[email] = guests[email] || []).push(h); covered[c] = true; }
    });
  }
  return {
    guests: Object.keys(guests).sort().map(e => ({ email: e, hubs: guests[e] })),
    missingHubs: hubs.filter(h => !covered[cleanHubName(h)])
  };
}

/* Daftar hub Hypercare v2 + calon guest */
function apiGetHypercareGuests(mode, week, topN, maxMetrics) {
  const m = mode === 'LM' ? 'LM' : 'FM';
  if (!cpRequireLead_(m)) return fail('Hanya City Lead/PIC.');
  const r = apiGetHypercareV2(m, week, topN, maxMetrics);
  if (!r || !r.success) return r || fail('Gagal memuat Hypercare v2.');
  const hubs = r.data.items.map(h => h.hub);
  const g = hcGuestsForHubs_(m, hubs);
  return ok({ hubs: hubs, guests: g.guests, missingHubs: g.missingHubs, organizer: Session.getActiveUser().getEmail() || '' });
}

function hcLogPlan_(row) {
  const ss = getSS();
  let sh = ss.getSheetByName(HC_PLAN_SHEET);
  if (!sh) {
    sh = ss.insertSheet(HC_PLAN_SHEET);
    sh.appendRow(['Plan_ID', 'Created_At', 'Created_By', 'Mode', 'Week', 'Hubs', 'Focus_Metrics', 'Title', 'Date', 'Start', 'Duration_Min', 'Guests', 'Meet_Link', 'Event_ID', 'Type', 'Message']);
    sh.getRange('A1:P1').setFontWeight('bold').setBackground('#f3f3f3');
    sh.setFrozenRows(1);
  }
  sh.appendRow(row);
  try { CacheService.getScriptCache().removeAll(['HC_PLAN_LATEST_FM', 'HC_PLAN_LATEST_LM']); } catch (e) {}
}

/* Buat event Google Calendar + Google Meet, catat rencana */
function apiCreateHypercareMeeting(p) {
  p = p || {};
  const m = p.mode === 'LM' ? 'LM' : 'FM';
  const access = getUserAccess(Session.getActiveUser().getEmail(), m);
  if (!access.isAllowed || !(access.role === 'City Lead' || access.role === 'PIC')) return fail('Hanya City Lead/PIC.');
  if (typeof Calendar === 'undefined') return fail('Layanan Google Calendar API belum diaktifkan di Apps Script (Services → Google Calendar API).');

  const title = String(p.title || '').trim().slice(0, 150);
  const date = String(p.date || ''), time = String(p.time || '');
  if (!title) return fail('Judul wajib diisi.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !/^\d{2}:\d{2}$/.test(time)) return fail('Tanggal/jam tidak valid.');
  const dur = Math.max(15, Math.min(180, parseInt(p.duration, 10) || 45));
  const start = new Date(date + 'T' + time + ':00+07:00');
  if (isNaN(start.getTime())) return fail('Tanggal/jam tidak valid.');
  if (start.getTime() < Date.now() - 5 * 60000) return fail('Waktu meeting sudah lewat.');
  const end = new Date(start.getTime() + dur * 60000);

  const emails = String(p.guests || '').split(/[\s,;]+/).map(s => s.toLowerCase().trim()).filter(Boolean);
  const bad = emails.filter(e => !hcIsEmail_(e));
  if (bad.length) return fail('Email tidak valid: ' + bad.slice(0, 3).join(', '));
  const uniq = Array.from(new Set(emails));
  if (!uniq.length) return fail('Minimal satu guest.');
  if (uniq.length > 50) return fail('Maksimal 50 guest.');

  const plan = apiGetHypercareV2(m, p.week, p.topN, p.maxMetrics);
  if (!plan || !plan.success) return plan || fail('Gagal memuat rencana Hypercare.');
  const hubs = plan.data.items.map(h => h.hub);
  const desc = String(p.description || '').split('(link Meet menyusul setelah meeting dibuat)').join('(lihat tombol Google Meet di undangan)').slice(0, 4000);

  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return fail('Server sedang sibuk, coba lagi.'); }
  try {
    const fmt = d => Utilities.formatDate(d, HC_TZ, "yyyy-MM-dd'T'HH:mm:ssXXX");
    const ev = Calendar.Events.insert({
      summary: title, description: desc,
      start: { dateTime: fmt(start), timeZone: HC_TZ }, end: { dateTime: fmt(end), timeZone: HC_TZ },
      attendees: uniq.map(e => ({ email: e })),
      conferenceData: { createRequest: { requestId: Utilities.getUuid(), conferenceSolutionKey: { type: 'hangoutsMeet' } } },
      guestsCanModify: false
    }, 'primary', { conferenceDataVersion: 1, sendUpdates: 'all' });

    const meet = ev.hangoutLink || ((ev.conferenceData && ev.conferenceData.entryPoints || []).filter(x => x.entryPointType === 'video')[0] || {}).uri || '';
    let warn = '';
    try {
      hcLogPlan_([Utilities.getUuid(), Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'), access.email, m, hcSafeCell_(plan.data.week),
        hcSafeCell_(hubs.join(', ')), hcSafeCell_(JSON.stringify(plan.data.items.map(h => ({ hub: h.hub, metrics: h.issues.map(i => i.key) })))),
        hcSafeCell_(title), date, time, dur, hcSafeCell_(uniq.join(', ')), meet, ev.id || '', 'Hypercare', hcSafeCell_(desc)]);
    } catch (e) { warn = 'Meeting dibuat, tetapi gagal mencatat ke sheet: ' + e.message; }
    return ok({ meetLink: meet, eventLink: ev.htmlLink || '', eventId: ev.id || '', guests: uniq.length, warning: warn });
  } catch (e) {
    return fail('Gagal membuat meeting: ' + e.message);
  } finally { lock.releaseLock(); }
}