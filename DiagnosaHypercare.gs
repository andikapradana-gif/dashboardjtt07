/************************************************************
 * DIAGNOSAHYPERCARE.GS — jalankan diagnosaHypercare() dari editor Apps Script, lalu kirim Execution log.
 * Hanya membaca, tidak mengubah data. Ubah MODE / WEEK di bawah sesuai yang Anda buka di dashboard
 * (WEEK persis seperti di filter periode, mis. '26 Oct - 01 Nov'; kosongkan '' untuk minggu terbaru).
 ************************************************************/
function diagnosaHypercare() {
  const MODE = 'LM';     // 'FM' atau 'LM'
  const WEEK = '';       // contoh: '26 Oct - 01 Nov'
  const TOPN = 10, MAXM = 3;

  const lap = (label, fn) => {
    const t = Date.now();
    let out;
    try { out = fn(); } catch (e) { Logger.log(label + ': ERROR ' + e.message); return null; }
    Logger.log(label + ': ' + (Date.now() - t) + ' ms');
    return out;
  };

  // 1) Sheet Hypercare_Input: ukuran & apakah cache-nya muat (batas CacheService 100 KB per kunci)
  const sh = getSS().getSheetByName('Hypercare_Input');
  if (sh) {
    Logger.log('Hypercare_Input: ' + sh.getLastRow() + ' baris x ' + sh.getLastColumn() + ' kolom (maks baris sheet: ' + sh.getMaxRows() + ')');
  } else {
    Logger.log('Hypercare_Input: sheet tidak ada');
  }
  CacheService.getScriptCache().remove('HYPERCARE_SAVED_INPUTS');
  const saved = lap('getHypercareSavedInputs (dingin)', () => getHypercareSavedInputs());
  if (saved) {
    const len = JSON.stringify(saved).length;
    Logger.log('Ukuran cache HYPERCARE_SAVED_INPUTS: ' + len + ' karakter → ' + (len > 100000 ? 'LEBIH dari 100.000, TIDAK bisa di-cache (dibaca ulang tiap buka menu)' : 'muat di cache') + ' | jumlah entri: ' + Object.keys(saved).length);
  }
  lap('getHypercareSavedInputs (hangat)', () => getHypercareSavedInputs());

  // 2) Bobot metrik
  lap('getMetricWeights', () => getMetricWeights());

  // 3) API lengkap seperti yang dipanggil browser + ukuran payload yang dikirim ke browser
  try { CacheService.getScriptCache().put('API_VER', String(Date.now()), 21600); } catch (e) {}
  const res = lap('apiGetHypercareV2 (dingin) ' + MODE + ' "' + WEEK + '"', () => apiGetHypercareV2(MODE, WEEK, TOPN, MAXM));
  if (res) {
    const size = JSON.stringify(res).length;
    Logger.log('success=' + res.success + (res.success ? '' : ' | ' + res.error) + ' | payload ke browser: ' + size + ' karakter' +
      (res.data ? ' | hub tampil: ' + (res.data.items || []).length + ' | minggu: ' + res.data.week : ''));
  }
  lap('apiGetHypercareV2 (hangat)', () => apiGetHypercareV2(MODE, WEEK, TOPN, MAXM));

  // 4) Antrean validasi yang dimuat setelah render
  const v = lap('apiGetHypercareValidation', () => apiGetHypercareValidation(MODE, res && res.data ? res.data.week : WEEK));
  if (v) Logger.log('validasi success=' + v.success + (v.success ? '' : ' | ' + v.error));
}