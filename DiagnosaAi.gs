/************************************************************
 * DIAGNOSAAI.GS — Alat diagnosa SEMENTARA untuk rekomendasi 5 Why (Gemini).
 * HANYA MEMBACA / LOGGING, tidak mengubah perilaku web. Aman dihapus setelah selesai.
 *
 * Cara pakai: Apps Script editor → pilih fungsi diagnosaGeminiAi → Run → buka "Execution log".
 * Key tidak ditampilkan (hanya panjangnya). Tempel hasil log ke chat.
 ************************************************************/
function diagnosaGeminiAi() {
  // 1. Ambil data konfigurasi dari Sheet "API"
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetAPI = ss.getSheetByName('API');
  
  if (!sheetAPI) {
    Logger.log('HASIL: Sheet bernama "API" tidak ditemukan di spreadsheet ini!');
    return;
  }

  // Membaca nilai dari Sheet "API"
  const provider = sheetAPI.getRange('B1').getValue();
  const apiKey = sheetAPI.getRange('B2').getValue();
  const model = sheetAPI.getRange('B3').getValue() || 'gemini-2.5-flash';

  Logger.log('provider=%s | model=%s | panjang key=%s', provider, model, (apiKey || '').toString().length);

  if (!apiKey) {
    Logger.log('HASIL: API key di sel B2 sheet "API" kosong!');
    return;
  }

  if (provider !== 'GEMINI') {
    Logger.log('Provider bukan GEMINI (saat ini: %s); diagnosa ini khusus Gemini.', provider);
    return;
  }

  // 2. Siapkan Context & Payload 5 Why
  const ctx = {
    mode: 'FM', hub: 'Hub Contoh', week: 'W1', metric: '(F1) Contoh Metrik', hint: '', cutoff: 75, value: 60, gap: 15,
    trend: [{ week: 'W-4', value: 80 }, { week: 'W-3', value: 78 }, { week: 'W-2', value: 70 }, { week: 'W-1', value: 66 }, { week: 'W1', value: 60 }],
    networkAvg: 80, rank: 2, hubCount: 10, hubsBelowCutoff: 2, others: [], prevRootCause: '', prevActionPlan: ''
  };

  // Gunakan fungsi prompt kustom jika ada, atau fallback ke teks tes sederhana
  let system = '', user = '';
  try { system = rc5SystemPrompt_(); } catch (e) { system = 'Anda adalah pakar analisis 5-Why.'; }
  try { user = rc5UserPrompt_(ctx); } catch (e) { user = 'Berikan contoh analisis 5-why sederhana format JSON.'; }

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent';

  const variants = [
    { nama: 'A. Seperti kode sekarang (maxOutputTokens 1200, thinking default)', gen: { temperature: 0.3, responseMimeType: 'application/json', maxOutputTokens: 1200 } },
    { nama: 'B. Pembanding (thinkingBudget 0, maxOutputTokens 2048)', gen: { temperature: 0.3, responseMimeType: 'application/json', maxOutputTokens: 2048, thinkingConfig: { thinkingBudget: 0 } } }
  ];

  // 3. Eksekusi Tes HTTP Fetch
  variants.forEach(function (v) {
    Logger.log('────── ' + v.nama);
    try {
      const res = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        muteHttpExceptions: true,
        headers: { 'x-goog-api-key': apiKey },
        payload: JSON.stringify({
          systemInstruction: { parts: [{ text: system }] },
          contents: [{ role: 'user', parts: [{ text: user }] }],
          generationConfig: v.gen
        })
      });

      const code = res.getResponseCode(), body = res.getContentText();
      Logger.log('HTTP %s', code);
      if (code < 200 || code >= 300) {
        Logger.log('Body error: ' + body.slice(0, 500));
        return;
      }
      const j = JSON.parse(body), c = (j.candidates || [])[0] || {};
      const text = ((c.content || {}).parts || []).map(function (p) { return p.text || ''; }).join('');
      Logger.log('finishReason=%s | promptFeedback=%s', c.finishReason, JSON.stringify(j.promptFeedback || null));
      Logger.log('usage=%s', JSON.stringify(j.usageMetadata || null));
      Logger.log('panjang teks jawaban=%s', text.length);
      Logger.log('awal jawaban: ' + text.slice(0, 200));

      try {
        if (typeof rc5Parse_ === 'function') {
          const r = rc5Parse_(text);
          Logger.log('PARSE: OK (%s why)', JSON.parse(r.whyChain).whys.length);
        } else {
          Logger.log('PARSE: Skip (rc5Parse_ tidak terdefinisi)');
        }
      } catch (pe) {
        Logger.log('PARSE: GAGAL → ' + pe.message);
      }
    } catch (e) {
      Logger.log('EXCEPTION: ' + e.message);
    }
  });
}