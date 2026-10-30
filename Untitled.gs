function debugMingguFM() {
  const raw = readSheetRaw(CONFIG.SHEETS.FM_REGULAR).rows;
  const norm = getFMRegular();
  Logger.log('Baris mentah: ' + raw.length + ' | lolos filter hub&tanggal: ' + norm.length);

  const baru = norm.filter(r => r.date >= '2026-10-26');
  Logger.log('Baris tgl >= 26 Okt (Reguler): ' + baru.length);
  const combo = {};
  baru.forEach(r => { const k = '[' + r.week + '] | bulan=' + r.month; combo[k] = (combo[k] || 0) + 1; });
  Logger.log('Kombinasi week|bulan: ' + JSON.stringify(combo));

  const high = getFMHigh().filter(r => r.date >= '2026-10-26');
  Logger.log('Baris tgl >= 26 Okt (High): ' + high.length);

  const tanpaWeek = norm.filter(r => !r.week).slice(-5).map(r => r.date);
  Logger.log('Contoh baris tanpa week: ' + JSON.stringify(tanpaWeek));
  Logger.log('getWeekList: ' + JSON.stringify(getWeekList('FM').slice(0, 5)));
}
function clearDataCache() {
  const c = CacheService.getScriptCache();
  clearBigCache_(c, ['DS_raw_FM_Reguler_Days', 'DS_raw_FM_High_Days']);
}