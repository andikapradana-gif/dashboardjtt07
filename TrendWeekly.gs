/************************************************************
 * TRENDWEEKLY.GS — Rekap Trend Weekly FM & LM ke sheet Rekapan_Trend_FM / _LM
 * (dipindah dari Scripts_BottomNav.html; ini kode server, bukan front-end)
 ************************************************************/

/**
 * Runner utama untuk mengeksekusi rekap Trend Weekly FM dan LM sekaligus
 * tanpa memicu timeout pada Google Apps Script.
 */
function runGenerateTrendWeeklyAll() {
  Logger.log('🚀 Memulai rekap Trend Weekly FM...');
  generateWeeklyTrendSheet('FM');
  
  Utilities.sleep(2000); // Jeda 2 detik untuk melepas lock memory spreadsheet
  
  Logger.log('🚀 Memulai rekap Trend Weekly LM...');
  generateWeeklyTrendSheet('LM');
  
  Logger.log('🎉 Seluruh proses rekap Trend Weekly (FM & LM) Selesai!');
}

/**
 * Generator utama Rekapan Trend Weekly (L3W, L2W, L1W, L0W) + Delta WoW
 * @param {string} mode - 'FM' atau 'LM'
 */
function generateWeeklyTrendSheet(mode) {
  const currentMode = mode || 'FM';
  const ss = getSS();
  const targetSheetName = `Rekapan_Trend_${currentMode}`;
  
  // 1. Ambil Data Mentah (FM Blended Weekly / LM [Backup NEW])
  const rawWeekly = currentMode === 'LM' ? getLMWeekly() : getFMWeekly();
  if (!rawWeekly || !rawWeekly.length) {
    Logger.log(`⚠️ Data mentah untuk mode ${currentMode} tidak ditemukan.`);
    return;
  }

  // Parser minggu: pakai parser yang sama dengan Hypercare/Checkpoint (tahun ikut disimpulkan, aman di pergantian tahun)
  function parseWeekToDate(weekStr) {
    if (!weekStr) return new Date(0);
    const d = hcv2WeekDate_(weekStr);
    if (d.getTime() > 0) return d;
    const numMatch = String(weekStr).match(/\d+/);   // format minggu berupa angka murni (38, 39, 40)
    return numMatch ? new Date(2000, 0, parseInt(numMatch[0], 10)) : new Date(0);
  }

  // 2. Kumpulkan Semua Minggu Unik & Petakan Tanggalnya
  const weeksMap = {};
  rawWeekly.forEach(r => { 
    if (r.week !== null && r.week !== undefined && r.week !== '') {
      const wStr = String(r.week).trim();
      if (!weeksMap[wStr]) {
        weeksMap[wStr] = parseWeekToDate(wStr);
      }
    }
  });

  // 3. Urutkan Minggu Kronologis (Tertua -> Terbaru) & Ambil 4 Minggu Terakhir (L3W s/d L0W)
  const sortedWeeks = Object.keys(weeksMap).sort((a, b) => weeksMap[a] - weeksMap[b]);
  const weeks = sortedWeeks.slice(-4); // Otomatis L3W, L2W, L1W, L0W (Termasuk October)

  if (!weeks.length) {
    Logger.log(`⚠️ Tidak ada minggu yang valid untuk mode ${currentMode}.`);
    return;
  }

  const metricsList = currentMode === 'LM' ? CONFIG.LM_METRICS : CONFIG.FM_METRICS;
  const weekSet = new Set(weeks);
  const agg = {};

  // 4. Aggregation Data per Metrik per Minggu (Single-Pass Engine)
  for (let i = 0; i < rawWeekly.length; i++) {
    const r = rawWeekly[i];
    const wStr = String(r.week).trim();
    if (!weekSet.has(wStr)) continue;

    for (let j = 0; j < metricsList.length; j++) {
      const m = metricsList[j];
      const v = r.metrics[m.key];
      if (v !== null && v !== undefined && !isNaN(v)) {
        const aggKey = `${wStr}_${m.key}`;
        if (!agg[aggKey]) agg[aggKey] = { sum: 0, count: 0 };
        agg[aggKey].sum += Number(v);
        agg[aggKey].count += 1;
      }
    }
  }

  // 5. Susun Header & Baris Output Tabel Statis
  const headers = ['Kode Metrik', 'Nama Metrik', ...weeks.map(w => `Week ${w}`), 'Delta (WoW)'];
  const outputRows = [headers];

  metricsList.forEach(m => {
    const row = [m.key, m.label];
    const scores = [];

    weeks.forEach(w => {
      const item = agg[`${w}_${m.key}`];
      const avgVal = (item && item.count > 0) ? (item.sum / item.count) : null;
      row.push(avgVal !== null ? (avgVal * 100).toFixed(2) + '%' : '-');
      scores.push(avgVal);
    });

    // Hitung Delta WoW: L0W (Terbaru/Kolom Kanan) dikurangi L1W (Minggu Sebelumnya)
    let delta = '-';
    if (scores.length >= 2) {
      const currentVal = scores[scores.length - 1]; // L0W
      const prevVal = scores[scores.length - 2];    // L1W
      if (currentVal !== null && prevVal !== null) {
        const dVal = (currentVal - prevVal) * 100;
        delta = (dVal > 0 ? '+' : '') + dVal.toFixed(2) + '%';
      }
    }
    row.push(delta);
    outputRows.push(row);
  });

  // 6. Batch Write ke Sheet Target (Value Only - Tanpa Rumus Heavy)
  let targetSheet = ss.getSheetByName(targetSheetName);
  if (!targetSheet) {
    targetSheet = ss.insertSheet(targetSheetName);
  } else {
    targetSheet.clearContents();
  }

  // Tulis seluruh tabel dalam 1 kali API call
  targetSheet.getRange(1, 1, outputRows.length, headers.length).setValues(outputRows);

  // Styling Header & Alignments
  const headerBg = currentMode === 'LM' ? '#2563eb' : '#f97316';
  targetSheet.getRange(1, 1, 1, headers.length)
             .setFontWeight('bold')
             .setBackground(headerBg)
             .setFontColor('#ffffff')
             .setHorizontalAlignment('center');

  targetSheet.getRange(2, 3, outputRows.length - 1, headers.length - 2)
             .setHorizontalAlignment('center');

  targetSheet.setFrozenRows(1);
  SpreadsheetApp.flush();
  
  Logger.log(`✅ Sheet ${targetSheetName} berhasil diperbarui! Week: ${weeks.join(' | ')}`);
}