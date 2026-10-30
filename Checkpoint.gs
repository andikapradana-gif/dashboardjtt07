/************************************************************
 * CHECKPOINT.GS — Google Meet checkpoint Jumat (guest = guest Hypercare)
 * Hanya City Lead / PIC. Memakai sheet Hypercare_Plan (lihat HypercareBlast.gs).
 ************************************************************/

/* Cek akses CL/PIC. Hasil di-cache 5 menit per email+mode (perubahan role di Access Web berlaku maks. 5 menit). */
function cpRequireLead_(mode) {
  const email = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  const cache = CacheService.getScriptCache();
  const ck = 'CP_LEAD_' + mode + '_' + (email ? hashKey_(email) : '-');
  if (email) {
    const c = cache.get(ck);
    if (c) { try { const o = JSON.parse(c); return o.lead ? o.a : null; } catch (e) {} }
  }
  const a = getUserAccess(email, mode);
  const isLead = !!(a && a.isAllowed && (a.role === 'City Lead' || a.role === 'PIC'));
  if (email) { try { cache.put(ck, JSON.stringify({ lead: isLead, a: isLead ? a : null }), 300); } catch (e) {} }
  return isLead ? a : null;
}

/* Jumat terdekat pada/setelah tanggal 'fromIso' (yyyy-MM-dd), tidak boleh sebelum hari ini */
function cpNextFriday_(fromIso) {
  const today = Utilities.formatDate(new Date(), HC_TZ, 'yyyy-MM-dd');
  let iso = /^\d{4}-\d{2}-\d{2}$/.test(fromIso || '') && fromIso > today ? fromIso : today;
  for (let i = 0; i < 8; i++) {
    const d = new Date(iso + 'T12:00:00+07:00');
    if (Utilities.formatDate(d, HC_TZ, 'u') === '5') return iso;
    iso = Utilities.formatDate(new Date(d.getTime() + 86400000), HC_TZ, 'yyyy-MM-dd');
  }
  return iso;
}

/* Rencana Hypercare terakhir untuk mode ini (di-cache 20 menit; dihapus saat rencana baru dicatat).
   Saat cache kosong hanya membaca ±60 baris terakhir & kolom A–O (tanpa kolom Message yang panjang). */
function cpLatestPlan_(mode) {
  const cache = CacheService.getScriptCache(), ck = 'HC_PLAN_LATEST_' + mode;
  const c = cache.get(ck);
  if (c) { try { const o = JSON.parse(c); if (o && o.none) return null; if (o && o.planId) return o; } catch (e) {} }
  let found = null;
  const sh = getSS().getSheetByName(HC_PLAN_SHEET);
  if (sh && sh.getLastRow() >= 2) {
    const lastRow = sh.getLastRow();
    const scan = (r1, r2) => {   // nomor baris sheet r1..r2, dipindai dari bawah ke atas
      if (r2 < r1) return null;
      const data = sh.getRange(r1, 1, r2 - r1 + 1, 15).getValues();
      for (let i = data.length - 1; i >= 0; i--) {
        const r = data[i];
        if (String(r[3]) === mode && String(r[14] || 'Hypercare') === 'Hypercare' && r[0]) {
          let focus = [];
          try { focus = JSON.parse(String(r[6] || '[]')); } catch (e) {}
          return {
            planId: String(r[0]), week: String(r[4] || ''), hubs: String(r[5] || ''),
            metrics: String(r[6] || ''), focus: focus, title: String(r[7] || ''), date: fmtDate(r[8]),
            guests: String(r[11] || ''), meetLink: String(r[12] || '')
          };
        }
      }
      return null;
    };
    const tail = Math.max(2, lastRow - 59);
    found = scan(tail, lastRow) || (tail > 2 ? scan(2, tail - 1) : null);
  }
  try { cache.put(ck, JSON.stringify(found || { none: true }), 1200); } catch (e) {}
  return found;
}

function apiGetCheckpointDraft(mode) {
  const m = mode === 'LM' ? 'LM' : 'FM';
  if (!cpRequireLead_(m)) return fail('Hanya City Lead/PIC.');
  const plan = cpLatestPlan_(m);
  if (!plan) return fail('Belum ada meeting Hypercare ' + m + '. Buat dulu lewat tombol "Blast & Google Meet".');
  return ok({ plan: plan, suggestedDate: cpNextFriday_(plan.date), organizer: Session.getActiveUser().getEmail() || '' });
}

function apiCreateCheckpointMeeting(p) {
  p = p || {};
  const m = p.mode === 'LM' ? 'LM' : 'FM';
  const access = cpRequireLead_(m);
  if (!access) return fail('Hanya City Lead/PIC.');
  if (typeof Calendar === 'undefined') return fail('Layanan Google Calendar API belum diaktifkan di Apps Script (Services → Google Calendar API).');

  const plan = cpLatestPlan_(m);
  if (!plan) return fail('Belum ada meeting Hypercare ' + m + '.');

  const title = String(p.title || '').trim().slice(0, 150);
  const date = String(p.date || ''), time = String(p.time || '');
  if (!title) return fail('Judul wajib diisi.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !/^\d{2}:\d{2}$/.test(time)) return fail('Tanggal/jam tidak valid.');
  const start = new Date(date + 'T' + time + ':00+07:00');
  if (isNaN(start.getTime())) return fail('Tanggal/jam tidak valid.');
  if (Utilities.formatDate(start, HC_TZ, 'u') !== '5') return fail('Checkpoint harus dijadwalkan pada hari Jumat.');
  if (start.getTime() < Date.now() - 5 * 60000) return fail('Waktu meeting sudah lewat.');
  const dur = Math.max(15, Math.min(180, parseInt(p.duration, 10) || 45));
  const end = new Date(start.getTime() + dur * 60000);

  const emails = String(p.guests || '').split(/[\s,;]+/).map(s => s.toLowerCase().trim()).filter(Boolean);
  const bad = emails.filter(e => !hcIsEmail_(e));
  if (bad.length) return fail('Email tidak valid: ' + bad.slice(0, 3).join(', '));
  const uniq = Array.from(new Set(emails));
  if (!uniq.length) return fail('Minimal satu guest.');
  if (uniq.length > 50) return fail('Maksimal 50 guest.');
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
      hcLogPlan_([Utilities.getUuid(), Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'), access.email, m,
        hcSafeCell_(plan.week), hcSafeCell_(plan.hubs), hcSafeCell_(plan.metrics), hcSafeCell_(title), date, time, dur,
        hcSafeCell_(uniq.join(', ')), meet, ev.id || '', 'Checkpoint', hcSafeCell_(desc)]);
      // kolom Q = ID rencana Hypercare induk (dipakai notulen di langkah berikutnya)
      const sh = getSS().getSheetByName(HC_PLAN_SHEET);
      if (!sh.getRange(1, 17).getValue()) sh.getRange(1, 17).setValue('Parent_Plan_ID').setFontWeight('bold').setBackground('#f3f3f3');
      sh.getRange(sh.getLastRow(), 17).setValue(plan.planId);
    } catch (e) { warn = 'Meeting dibuat, tetapi gagal mencatat ke sheet: ' + e.message; }
    return ok({ meetLink: meet, eventLink: ev.htmlLink || '', guests: uniq.length, warning: warn });
  } catch (e) {
    return fail('Gagal membuat meeting: ' + e.message);
  } finally { lock.releaseLock(); }
}