/************************************************************
 * DIAGNOSA.GS — jalankan diagnosaLemot() dari editor Apps Script (Run), lalu buka View > Logs / Execution log.
 * Tidak mengubah data apa pun. Hanya membaca sheet dan mencatat waktu.
 ************************************************************/

function diagColLetter_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function diagnosaLemot() {
  const ss = getSS();
  const isTransient = v => typeof v === 'string' && (v === 'Loading...' || v === 'Loading…' || /^#(N\/A|REF!|VALUE!|ERROR!|DIV\/0!|NAME\?)/.test(v));

  [CONFIG.SHEETS.FM_WEEKLY, CONFIG.SHEETS.LM_BACKUP].forEach(function (name) {
    const sh = ss.getSheetByName(name);
    if (!sh) { Logger.log('[' + name + '] sheet tidak ditemukan'); return; }
    const lr = sh.getLastRow(), lc = sh.getLastColumn();

    let t = Date.now();
    const all = sh.getRange(1, 1, lr, lc).getValues();
    const readMs = Date.now() - t;

    t = Date.now();
    const json = JSON.stringify({ headers: all[0], rows: all.slice(1) });
    const jsonMs = Date.now() - t;

    let transientCount = 0, first = '';
    for (let r = 1; r < all.length; r++) {
      for (let c = 0; c < all[r].length; c++) {
        if (isTransient(all[r][c])) {
          transientCount++;
          if (!first) first = diagColLetter_(c + 1) + (r + 1) + ' = "' + all[r][c] + '"';
        }
      }
    }

    Logger.log('[' + name + '] ' + lr + ' baris x ' + lc + ' kolom | baca sheet: ' + readMs + ' ms | JSON: ' +
      json.length + ' karakter (' + jsonMs + ' ms) | batas cache 1.500.000 → ' +
      (json.length > 1500000 ? 'TIDAK AKAN DI-CACHE (kebesaran)' : 'muat di cache'));
    Logger.log('[' + name + '] sel error/Loading: ' + transientCount +
      (transientCount ? ' (pertama: ' + first + ') → selama ada ini, sheet TIDAK di-cache dan dibaca ulang tiap permintaan' : ' → aman'));
  });

  // Waktu eksekusi API yang dipakai menu Trend Mingguan & Hypercare (cache API dibuang dulu agar yang diukur kondisi dingin)
  try { CacheService.getScriptCache().put('API_VER', String(Date.now()), 21600); } catch (e) {}
  ['FM', 'LM'].forEach(function (mode) {
    let t = Date.now();
    const a = apiGetComprehensiveTrend(mode, 'ALL_HUB', '', 'Semua');
    Logger.log('[' + mode + '] apiGetComprehensiveTrend ALL_HUB: ' + (Date.now() - t) + ' ms | success=' + (a && a.success) + (a && !a.success ? ' | ' + a.error : ''));
    t = Date.now();
    const h = apiGetHypercareV2(mode, '', 10, 3);
    Logger.log('[' + mode + '] apiGetHypercareV2: ' + (Date.now() - t) + ' ms | success=' + (h && h.success) + (h && !h.success ? ' | ' + h.error : ''));
  });
}