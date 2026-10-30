/************************************************************
 * HYPERCAREVALIDASI.GS — Temuan lapangan, jenis masalah (Plan / Eksekusi) & validasi
 *
 * Alur: Hub Lead isi temuan + bukti → jawab pertanyaan penentu → status "Menunggu validasi"
 *       → City Lead / PIC: Setujui | Ubah jenis | Minta bukti.
 *
 * Penyimpanan: dua kolom tambahan di sheet "Hypercare_Input"
 *   M = Temuan_JSON   {desc, ev:[{t,v}], q, jenis, faktor:[...]}   (tanpa template; data lama yang punya 'tpl' diabaikan)
 *   N = Validasi_JSON {state, by, at, note, from, to}
 * Kolom A–L tidak berubah, jadi form lama tetap jalan.
 ************************************************************/

const HCV_PENDING_ = 'Menunggu validasi';
const HCV_EVTYPES_ = ['Link foto', 'Jam kejadian', 'Data'];
const HCV_FACTORS_ = ['Sistem/IT', 'SDM', 'Fasilitas & kendaraan', 'Lainnya'];   // boleh lebih dari satu

function hcvParse_(s) { try { const o = s ? JSON.parse(s) : null; return (o && typeof o === 'object') ? o : {}; } catch (e) { return {}; } }
function hcvClip_(s, n) { return String(s === null || s === undefined ? '' : s).replace(/\s+/g, ' ').trim().slice(0, n); }
function hcvNorm_(s) { return String(s || '').toLowerCase().replace(/[^a-z0-9]/g, ''); }
function hcvNow_() { return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm'); }
function hcvModeOf_(metricKey) { return String(metricKey || '').indexOf('L_') === 0 ? 'LM' : 'FM'; }
function hcvIsFull_(a) { return a.role === 'City Lead' || a.role === 'PIC' || (a.hubAccess || []).indexOf('Semua') >= 0; }
function hcvIsValidator_(a) { return a.isAllowed && (a.role === 'City Lead' || a.role === 'PIC'); }

function hcvSheet_() {
  const ss = getSS();
  let sheet = ss.getSheetByName('Hypercare_Input');
  if (!sheet) {
    sheet = ss.insertSheet('Hypercare_Input');
    sheet.appendRow(['Key_ID', 'Hub', 'Week', 'Metric_Key', 'Root_Cause', 'Action_Plan', 'Updated_By', 'Updated_At', 'Status', 'Due_Date', 'PIC', 'Why_Chain']);
    sheet.setFrozenRows(1);
  }
  sheet.getRange(1, 13, 1, 2).setValues([['Temuan_JSON', 'Validasi_JSON']]);
  sheet.getRange(1, 1, 1, 14).setFontWeight('bold').setBackground('#f3f3f3');
  return sheet;
}

/* Cari baris dengan aturan pencocokan yang sama seperti saveHypercareInputUnsafe_ */
function hcvFindRow_(sheet, hub, week, metricKey) {
  const keyId = buildHypercareKey(hub, week, metricKey);
  const cH = cleanHubName(hub), cW = hcvNorm_(week), cM = String(metricKey || '').toLowerCase().trim();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase().trim() === keyId ||
        (cleanHubName(data[i][1]) === cH && hcvNorm_(data[i][2]) === cW && String(data[i][3]).toLowerCase().trim() === cM)) return i + 1;
  }
  return -1;
}

function hcvJenisFromQ_(q) { return q === 'Ya' ? 'Eksekusi' : (q === 'Tidak' ? 'Plan' : ''); }

/* ── Hub Lead: simpan temuan, bukti, jawaban penentu, faktor ── */
function apiSaveHypercareFinding(hub, week, metricKey, f) {
  try {
    const access = getUserAccess(Session.getActiveUser().getEmail(), hcvModeOf_(metricKey));
    if (!access.isAllowed) return fail('Akses ditolak.');
    if (access.hubAccess && access.hubAccess.indexOf('Semua') < 0 &&
        !access.hubAccess.some(h => cleanHubName(h) === cleanHubName(hub))) return fail('Anda tidak memiliki akses untuk hub ini.');
    f = f || {};
    const ev = (Array.isArray(f.ev) ? f.ev : []).slice(0, 3)
      .map(e => ({ t: HCV_EVTYPES_.indexOf(e && e.t) >= 0 ? e.t : '', v: hcvClip_(e && e.v, 300) }))
      .filter(e => e.t && e.v);
    const clean = {
      desc: hcvClip_(f.desc, 500),
      ev: ev,
      q: (f.q === 'Ya' || f.q === 'Tidak') ? f.q : '',
      faktor: (Array.isArray(f.faktor) ? f.faktor : []).filter((x, i, a) => HCV_FACTORS_.indexOf(x) >= 0 && a.indexOf(x) === i)
    };

    const lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (e) { return fail('Server sedang sibuk, coba lagi sebentar.'); }
    try {
      const sheet = hcvSheet_();
      let row = hcvFindRow_(sheet, hub, week, metricKey);
      const email = Session.getActiveUser().getEmail() || 'Anonymous';
      if (row < 0) {
        sheet.appendRow([buildHypercareKey(hub, week, metricKey), hcvClip_(hub, 100), hcvClip_(week, 50), hcvClip_(metricKey, 20), '', '', email, hcvNow_(), 'Open', '', '', '', '', '']);
        row = sheet.getLastRow();
      }
      const old = hcvParse_(sheet.getRange(row, 13).getValue());
      const val = hcvParse_(sheet.getRange(row, 14).getValue());
      // Jenis hanya dihitung ulang bila jawaban penentu berubah; koreksi City Lead/PIC tidak tertimpa
      clean.jenis = (clean.q !== (old.q || '')) ? hcvJenisFromQ_(clean.q) : (old.jenis || hcvJenisFromQ_(clean.q));
      sheet.getRange(row, 13).setValue(JSON.stringify(clean));
      // Bukti atau jenis berubah setelah disetujui → validasi harus diulang
      const changed = JSON.stringify(old.ev || []) !== JSON.stringify(clean.ev) || (old.jenis || '') !== clean.jenis;
      if (val.state === 'Disetujui' && changed) sheet.getRange(row, 14).setValue('');
      CacheService.getScriptCache().remove('HYPERCARE_SAVED_INPUTS');
      return ok({ jenis: clean.jenis });
    } finally { lock.releaseLock(); }
  } catch (e) { return fail(e.message); }
}

/* Dipanggil apiSaveHypercareInput sebelum status disimpan. Mengembalikan pesan error atau '' */
function hcvCheckStatus_(access, hub, week, metricKey, status) {
  if (status !== HCV_PENDING_ && status !== 'Done') return '';
  const s = getHypercareSavedInputs()[buildHypercareKey(hub, week, metricKey)] || {};
  const t = s.temuan || {};
  if (status === HCV_PENDING_) {
    if (!(t.ev && t.ev.length) || !t.jenis) return 'Sebelum dikirim untuk validasi: tambahkan minimal satu bukti dan jawab pertanyaan penentu.';
    return '';
  }
  // Done: Hub Lead hanya boleh bila sudah disetujui City Lead / PIC (atau memang sudah Done sebelumnya)
  if (!hcvIsFull_(access) && s.status !== 'Done' && !(s.val && s.val.state === 'Disetujui')) return 'Status Done menunggu persetujuan City Lead / PIC. Pilih "Menunggu validasi".';
  return '';
}

/* ── City Lead / PIC: ringkasan Plan vs Eksekusi + antrean validasi ── */
function apiGetHypercareValidation(mode, week) {
  try {
    const m = mode === 'LM' ? 'LM' : 'FM';
    const access = getUserAccess(Session.getActiveUser().getEmail(), m);
    if (!hcvIsValidator_(access)) return fail('Hanya City Lead dan PIC.');
    const metrics = m === 'LM' ? CONFIG.LM_METRICS : CONFIG.FM_METRICS;
    const label = {}; metrics.forEach(x => { label[x.key] = x.label; });
    const sheet = getSS().getSheetByName('Hypercare_Input');
    const out = { week: week, summary: { total: 0, plan: 0, eksekusi: 0, validated: 0, pending: 0, faktor: {} }, queue: [] };
    if (!sheet) return ok(out);
    const wk = hcvNorm_(week);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const r = data[i], key = String(r[3] || '');
      if (!r[0] || (hcvModeOf_(key) !== m) || (wk && hcvNorm_(r[2]) !== wk)) continue;
      const t = hcvParse_(r[12]), v = hcvParse_(r[13]), status = r[8] || 'Open';
      if (t.jenis) {
        out.summary.total++;
        if (t.jenis === 'Plan') out.summary.plan++; else out.summary.eksekusi++;
        if (v.state === 'Disetujui') out.summary.validated++;
        [].concat(t.faktor || []).forEach(x => { out.summary.faktor[x] = (out.summary.faktor[x] || 0) + 1; });
      }
      if (status === HCV_PENDING_) {
        out.summary.pending++;
        out.queue.push({ hub: r[1], week: r[2], key: key, metric: label[key] || key, temuan: t, val: v,
          rootCause: r[4] || '', by: String(r[6] || '').split('@')[0], at: r[7] ? Utilities.formatDate(new Date(r[7]), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm') : '' });
      }
    }
    out.queue.sort((a, b) => ((a.temuan.ev || []).length ? 1 : 0) - ((b.temuan.ev || []).length ? 1 : 0));   // tanpa bukti di atas
    return ok(out);
  } catch (e) { return fail(e.message); }
}

/* decision: 'setuju' | 'ubah' | 'bukti' */
function apiValidateHypercare(hub, week, metricKey, decision, note) {
  try {
    const access = getUserAccess(Session.getActiveUser().getEmail(), hcvModeOf_(metricKey));
    if (!hcvIsValidator_(access)) return fail('Hanya City Lead dan PIC yang dapat memvalidasi.');
    if (['setuju', 'ubah', 'bukti'].indexOf(decision) < 0) return fail('Keputusan tidak valid.');
    note = hcvClip_(note, 300);
    if (decision !== 'setuju' && note.length < 5) return fail('Isi alasan singkat (minimal 5 karakter).');

    const lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (e) { return fail('Server sedang sibuk, coba lagi sebentar.'); }
    try {
      const sheet = hcvSheet_();
      const row = hcvFindRow_(sheet, hub, week, metricKey);
      if (row < 0) return fail('Data tidak ditemukan.');
      if ((sheet.getRange(row, 9).getValue() || 'Open') !== HCV_PENDING_) return fail('Item ini tidak sedang menunggu validasi (mungkin sudah diproses).');
      const t = hcvParse_(sheet.getRange(row, 13).getValue());
      const email = Session.getActiveUser().getEmail() || 'Anonymous';
      const v = { by: email, at: hcvNow_(), note: note };
      let newStatus;
      if (decision === 'setuju') {
        if (!(t.ev && t.ev.length) || !t.jenis) return fail('Tidak bisa disetujui: bukti atau jenis belum lengkap. Pilih "Minta bukti".');
        v.state = 'Disetujui'; newStatus = 'Done';
      } else if (decision === 'ubah') {
        v.state = 'Dikembalikan'; v.from = t.jenis || ''; v.to = t.jenis === 'Plan' ? 'Eksekusi' : 'Plan';
        t.jenis = v.to; newStatus = 'In progress';
        sheet.getRange(row, 13).setValue(JSON.stringify(t));
      } else {
        v.state = 'Minta bukti'; newStatus = 'In progress';
      }
      sheet.getRange(row, 9).setValue(newStatus);
      sheet.getRange(row, 14).setValue(JSON.stringify(v));
      CacheService.getScriptCache().remove('HYPERCARE_SAVED_INPUTS');
      return ok({ state: v.state, status: newStatus });
    } finally { lock.releaseLock(); }
  } catch (e) { return fail(e.message); }
}
