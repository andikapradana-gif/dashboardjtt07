/************************************************************
 * SHARINGBLAST.GS — Guest & Google Meet untuk Sharing Session
 * Hanya City Lead / PIC. Memakai helper dari HypercareBlast, Checkpoint, SharingSession.
 ************************************************************/

function sbScope_(s) { return (s === 'FM' || s === 'LM') ? s : 'ALL'; }
function sbModes_(scope) { return scope === 'FM' ? ['FM'] : (scope === 'LM' ? ['LM'] : ['FM', 'LM']); }

/* Mode dalam cakupan yang boleh diakses user (harus CL/PIC) */
function sbAllowed_(scope) {
  const modes = []; let access = null;
  sbModes_(scope).forEach(m => {
    const a = cpRequireLead_(m);
    if (a) { modes.push(m); if (!access) access = a; }
  });
  return { modes: modes, access: access };
}

/* Guest Sharing Session = SEMUA Hub Lead + City Lead + PIC di sheet Access Web (email valid, tanpa duplikat).
   Tidak difilter berdasarkan hub/data. Bila satu email muncul di beberapa baris,
   role yang dipakai: City Lead > PIC > Hub Lead. */
function sbGuests_(modes) {
  const data = getAccessRows_() || [];
  const clean = s => String(s || '').trim();
  const validHub = h => h && h !== '-' && h.toLowerCase() !== 'null';
  const RANK = { 'City Lead': 3, 'PIC': 2, 'Hub Lead': 1 };
  const roleOf = r => /city\s*[-_]?\s*lead/i.test(r) ? 'City Lead'
                    : /hub\s*[-_]?\s*lead/i.test(r) ? 'Hub Lead'
                    : /\bpic\b/i.test(r) ? 'PIC' : '';
  const out = {}, roleSeen = {};

  for (let i = 1; i < data.length; i++) {
    const email = clean(data[i][1]).toLowerCase();
    if (!hcIsEmail_(email)) continue;
    const rawRole = clean(data[i][2]);
    roleSeen[rawRole] = (roleSeen[rawRole] || 0) + 1;
    const role = roleOf(rawRole);
    if (!role) continue;

    const o = out[email] = out[email] || { email: email, role: role, hubs: [] };
    if (RANK[role] > RANK[o.role]) o.role = role;

    if (role === 'Hub Lead') {
      modes.forEach(m => {
        const list = m === 'LM' ? [data[i][4], data[i][5]] : String(data[i][3] || '').split(',');
        list.map(clean).filter(validHub).forEach(h => { if (o.hubs.indexOf(h) < 0) o.hubs.push(h); });
      });
    }
  }
  Logger.log('Sharing guests, role di Access Web: ' + JSON.stringify(roleSeen));

  const guests = Object.keys(out).sort().map(e => out[e]);
  const count = r => guests.filter(g => g.role === r).length;
  return { guests: guests, hubLeads: count('Hub Lead'), cityLeads: count('City Lead'), pics: count('PIC') };
}

function apiGetSharingGuests(scope) {
  try {
    const sc = sbScope_(scope);
    const g = sbAllowed_(sc);
    if (!g.modes.length) return fail('Hanya City Lead/PIC.');
    const r = sbGuests_(g.modes);
return ok({
  guests: r.guests, hubLeads: r.hubLeads, cityLeads: r.cityLeads, pics: r.pics,
  modes: g.modes,
  organizer: Session.getActiveUser().getEmail() || ''
});
  } catch (e) { return fail(e.message); }
}

/* Buat event Google Calendar + Meet untuk Sharing Session, lalu catat ke Hypercare_Plan (Type = Sharing) */
function apiCreateSharingMeeting(p) {
  p = p || {};
  const sc = sbScope_(p.scope);
  const g = sbAllowed_(sc);
  if (!g.modes.length) return fail('Hanya City Lead/PIC.');
  if (typeof Calendar === 'undefined') return fail('Layanan Google Calendar API belum diaktifkan di Apps Script (Services → Google Calendar API).');

  const title = String(p.title || '').trim().slice(0, 150);
  const date = String(p.date || ''), time = String(p.time || '');
  if (!title) return fail('Judul wajib diisi.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !/^\d{2}:\d{2}$/.test(time)) return fail('Tanggal/jam tidak valid.');
  const dur = Math.max(15, Math.min(180, parseInt(p.duration, 10) || 60));
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

  // narasumber dihitung ulang di server (cache withApiCache), bukan dari kiriman browser
  const compare = p.compare === 'L2_L1' ? 'L2_L1' : 'L1_L0';
  const picks = [];
  g.modes.forEach(m => {
    const r = ssComputeMode_(m, String(p.week || ''), compare, p.criteria);
    if (r && r.success && r.data.pick) {
      picks.push({ mode: m, hub: r.data.pick.hub, week: r.data.week, metrics: r.data.pick.improved.map(i => i.key) });
    }
  });
  if (!picks.length) return fail('Belum ada rekomendasi narasumber untuk dijadwalkan.');

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
      hcLogPlan_([Utilities.getUuid(), Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'), g.access.email, sc,
        hcSafeCell_(picks[0].week),
        hcSafeCell_(picks.map(x => x.mode + ': ' + x.hub).join(', ')),
        hcSafeCell_(JSON.stringify(picks.map(x => ({ mode: x.mode, hub: x.hub, metrics: x.metrics })))),
        hcSafeCell_(title), date, time, dur, hcSafeCell_(uniq.join(', ')), meet, ev.id || '', 'Sharing', hcSafeCell_(desc)]);
    } catch (e) { warn = 'Meeting dibuat, tetapi gagal mencatat ke sheet: ' + e.message; }
    return ok({ meetLink: meet, eventLink: ev.htmlLink || '', eventId: ev.id || '', guests: uniq.length, warning: warn });
  } catch (e) {
    return fail('Gagal membuat meeting: ' + e.message);
  } finally { lock.releaseLock(); }
}