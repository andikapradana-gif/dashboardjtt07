/************************************************************
 * SHARINGSESSION.GS — Rekomendasi narasumber sharing session
 * 1 hub FM + 1 hub LM dengan kenaikan skor mingguan (WoW) terbaik.
 * Hanya City Lead / PIC. Memakai helper dari Code, HypercareV2, Checkpoint.
 *
 * Logika:
 *  - Skor hub = kolom SCORE weekly (hanya baris yang lolos rowScoreUsable_).
 *  - Kandidat = hub yang skornya NAIK minimal SS_MIN_DELTA poin dari minggu sebelumnya
 *    dan punya minimal 1 metrik yang naik.
 *  - Urutan: kenaikan skor terbesar → naik berturut-turut terlama → skor saat ini tertinggi.
 *  - Metrik kunci = metrik yang naik, diurutkan menurut kontribusi ke skor hub
 *    (bobot × kenaikan ÷ total bobot metrik yang terisi).
 *  - Bila hub pernah mengisi root cause / action plan Hypercare pada minggu pembanding,
 *    catatannya ikut ditampilkan sebagai bahan cerita.
 ************************************************************/

const SS_MIN_DELTA = 0.05;   // poin; kenaikan lebih kecil dianggap flat
const SS_ALT_COUNT = 2;      // jumlah hub alternatif per divisi

function ssWeekIndex_(all) {
  const set = {};
  all.forEach(r => { set[String(r.week).trim()] = true; });
  return Object.keys(set).sort((a, b) => hcv2WeekDate_(a) - hcv2WeekDate_(b));
}

const SS_CRITERIA = ['delta', 'streak', 'score', 'crossed'];
/* Pembanding: 'L1_L0' = minggu terpilih vs sebelumnya; 'L2_L1' = satu minggu lebih awal vs sebelumnya.
   Kriteria urutan: delta (kenaikan terbesar), streak (naik berturut-turut terlama), score (skor tertinggi), crossed (metrik lolos cutoff terbanyak). */
function ssComputeMode_(m, weekFilter, compare, criteria) {
  const cutoff = getHypercareCutoffFor_(m);
  const off = compare === 'L2_L1' ? 1 : 0;
  const crit = SS_CRITERIA.indexOf(criteria) >= 0 ? criteria : 'delta';
  const cacheKey = 'sharing_v2_' + m + '_' + (weekFilter || '-') + '_' + off + '_' + crit + '_th_' + cutoff;
  return withApiCache(cacheKey, function () {
    try {
      const metricsList = m === 'LM' ? CONFIG.LM_METRICS : CONFIG.FM_METRICS;
      const all = (m === 'LM' ? getLMWeekly() : getFMWeekly()).filter(r => r.hub && r.week && rowHasMetricData_(r));
      if (!all.length) return fail('Belum ada data mingguan yang terisi.');

      const weeks = ssWeekIndex_(all);
      let week = String(weekFilter || '').trim(), fallback = false;
      if (!week || weeks.indexOf(week) < 0) { fallback = !!week; week = weeks[weeks.length - 1]; }
      const idx = weeks.indexOf(week) - off;
      if (idx < 1) return fail('Belum ada minggu pembanding untuk ' + (off ? 'L2W → L1W' : 'minggu ' + week) + '.');
      week = weeks[idx];
      const prevWeek = weeks[idx - 1];

      const byWeek = {};
      weeks.forEach(w => { byWeek[w] = {}; });
      all.forEach(r => { byWeek[String(r.week).trim()][cleanHubName(r.hub)] = r; });

      const pctOf = v => (v === null || v === undefined || isNaN(v)) ? null : scoreToPct_(v);
      const scoreAt = (w, ch) => { const r = byWeek[w] && byWeek[w][ch]; return rowScoreUsable_(r) ? scoreToPct_(r.score) : null; };

      const wInfo = hcv2ResolveWeights_(m, metricsList);
      const cur = byWeek[week], prev = byWeek[prevWeek];

      const cands = [];
      let deltaSum = 0, deltaN = 0, upCount = 0;
      Object.keys(cur).forEach(ch => {
        const rc = cur[ch], rp = prev[ch];
        if (!rp) return;
        const sc = scoreAt(week, ch), sp = scoreAt(prevWeek, ch);
        if (sc === null || sp === null) return;
        const delta = sc - sp;
        deltaSum += delta; deltaN++;
        if (delta > SS_MIN_DELTA) upCount++;

        // berapa minggu berturut-turut skor naik (maks 4 minggu ke belakang)
        let streak = 0;
        for (let i = idx; i >= 1 && streak < 4; i--) {
          const a = scoreAt(weeks[i], ch), b = scoreAt(weeks[i - 1], ch);
          if (a !== null && b !== null && a - b > SS_MIN_DELTA) streak++; else break;
        }

        let availW = 0;
        metricsList.forEach(mt => { if (pctOf(rc.metrics[mt.key]) !== null) availW += wInfo.map[mt.key]; });
        if (!availW) return;

        const improved = [], below = [];
        metricsList.forEach(mt => {
          const vc = pctOf(rc.metrics[mt.key]), vp = pctOf(rp.metrics[mt.key]);
          const w = wInfo.map[mt.key];
          if (vc !== null && vp !== null && vc - vp > 0.005) {
            improved.push({ key: mt.key, label: mt.label, prev: vp, cur: vc, delta: vc - vp, contribution: w * (vc - vp) / availW, crossed: vp < cutoff && vc >= cutoff });
          }
          if (vc !== null && vc < cutoff) below.push({ key: mt.key, label: mt.label, cur: vc, gap: cutoff - vc, impact: w * (cutoff - vc) / availW });
        });
        const crossedCount = improved.filter(x => x.crossed).length;
        improved.sort((a, b) => b.contribution - a.contribution);
        below.sort((a, b) => b.impact - a.impact);

        cands.push({ hub: rc.hub, score: sc, prevScore: sp, delta: delta, streak: streak, crossedCount: crossedCount, improved: improved.slice(0, 3), improvedCount: improved.length, below: below.slice(0, 2), belowCount: below.length });
      });

      const sorters = {
        delta:   (a, b) => b.delta - a.delta || b.streak - a.streak || b.score - a.score,
        streak:  (a, b) => b.streak - a.streak || b.delta - a.delta || b.score - a.score,
        score:   (a, b) => b.score - a.score || b.delta - a.delta,
        crossed: (a, b) => b.crossedCount - a.crossedCount || b.delta - a.delta || b.score - a.score
      };
      const eligible = cands.filter(c => c.delta > SS_MIN_DELTA && c.improvedCount > 0 && (crit !== 'crossed' || c.crossedCount > 0))
        .sort(sorters[crit]);
      eligible.forEach((c, i) => { c.rank = i + 1; });

      return ok({
        mode: m, week: week, prevWeek: prevWeek, weekFallback: fallback, cutoff: cutoff, compare: off ? 'L2_L1' : 'L1_L0', criteria: crit,
        hubCount: deltaN, hubsUp: upCount, avgDelta: deltaN ? deltaSum / deltaN : null,
        pick: eligible.length ? eligible[0] : null,
        alternates: eligible.slice(1, 1 + SS_ALT_COUNT).map(c => ({ hub: c.hub, delta: c.delta, score: c.score, streak: c.streak }))
      });
    } catch (e) { return fail(e.message); }
  });
}

/* Catatan Hypercare & email narasumber digabung SETELAH cache agar selalu terbaru */
function ssAttachExtras_(d, m) {
  if (!d || !d.pick) return;
  try {
    const saved = getHypercareSavedInputs();
    d.pick.improved.forEach(i => {
      const s = saved[buildHypercareKey(d.pick.hub, d.prevWeek, i.key)] || saved[buildHypercareKey(d.pick.hub, d.week, i.key)];
      if (s && (s.rootCause || s.actionPlan)) i.note = { rootCause: s.rootCause || '', actionPlan: s.actionPlan || '', pic: s.pic || '', status: s.status || '' };
    });
  } catch (e) {}
  try { d.pick.speakers = hcGuestsForHubs_(m, [d.pick.hub]).guests.map(g => g.email); } catch (e) { d.pick.speakers = []; }
}

function apiGetSharingSession(weekFilter, compare, criteria) {
  try {
    const res = {};
    let any = false;
    ['FM', 'LM'].forEach(m => {
      if (!cpRequireLead_(m)) { res[m] = { denied: true }; return; }
      any = true;
      const r = ssComputeMode_(m, weekFilter, compare === 'L2_L1' ? 'L2_L1' : 'L1_L0', criteria);
      if (r && r.success) { ssAttachExtras_(r.data, m); res[m] = r.data; }
      else res[m] = { error: r ? r.error : 'Gagal menghitung.' };
    });
    if (!any) return fail('Menu Sharing Session hanya untuk City Lead dan PIC.');
    return ok(res);
  } catch (e) { return fail(e.message); }
}
