/************************************************************
 * CONFIG.GS — Unified FM & LM Configuration · JTT 07
 ************************************************************/

const CONFIG = {
  SPREADSHEET_ID: '1w-LqqOy60X6vapP5wzFrrHKxQi_cbuHsSx7iie10UR0',
  
  SHEETS: {
    FM_REGULAR: 'FM Reguler Days',
    FM_HIGH:    'FM High Days',
    FM_WEEKLY:  'FM Blended Weekly',
    FM_MONTHLY: 'FM Blended Monthly',
    LM_BACKUP:  'LM [Backup NEW]',
    ACCESS_WEB: 'Access Web',
    WEIGHTS:    'Matriks & Bobot'
  },
  
  TIMEZONE: 'GMT+7',

  // Satu sumber target KPI (dipakai backend & dikirim ke frontend)
  TARGET_PCT: 99.90,
  WARN_PCT: 99.00,
  
  THRESHOLD: {
    TERCAPAI: 0.999
  },
  
  // Metrik FM (9 Metrik)
  FM_METRICS: [
    { key: 'A1', label: '(A1) Sameday Pickup (ADO)' },
    { key: 'A2', label: '(A2) Sameday Pickup (Seller)' },
    { key: 'A3', label: '(A3) Pickup Aging > 48hr' },
    { key: 'A4', label: '(A4) Bulky Compliance' },
    { key: 'A5', label: '(A5) Stuck Parcel > 24hr' },
    { key: 'B1', label: '(B1) Clean Floor before 00' },
    { key: 'B2', label: '(B2) Clean Floor after 00' },
    { key: 'B4', label: '(B3) SLA Packed to Returned' },
    { key: 'C1', label: '(C1) Daily Active Operator' }
  ],

  // Metrik LM (9 Metrik)
  LM_METRICS: [
    { key: 'L_A1', label: '(A1) %couriers with penalty point >= x vs target' },
    { key: 'L_B2', label: '(B2) 2W Dedicated & Plus Ideal Daily Active' },
    { key: 'L_C1', label: '(C1) 2W Dedicated & Plus Monthly Productivity' },
    { key: 'L_C2', label: '(C2) 3W & 4W Dedicated Productivity' },
    { key: 'L_C3', label: '(C3) LM Hub operator Productivity vs target' },
    { key: 'L_D1', label: '(D1) %Sameday Delivered STD' },
    { key: 'L_D2', label: '(D2) %NDD ECO Compliance' },
    { key: 'L_D3', label: '(D3) % Aging D+1 Parcels Delivered (Expedite Pickup)' },
    { key: 'L_D4', label: '(D4) % Return Rate' }
  ],
  
  FM_COLUMNS: {
    DAILY:   { DATE: 0, WEEK: 1, MONTH: 2, HUB: 3, METRICS_START: 4, METRICS_COUNT: 9 },
    // SCORE_START: skor per metrik "FM Blended Weekly" kolom N:V (indeks 13–21), urutan sama dengan FM_METRICS
    WEEKLY:  { WEEK: 0, MONTH: 1, HUB: 2, SCORE: 22, METRICS_START: 3, METRICS_COUNT: 9, SCORE_START: 13 },
    MONTHLY: { MONTH: 0, HUB: 1, SCORE: 21, METRICS_START: 2, METRICS_COUNT: 9 }
  },

  LM_COLUMNS: {
    DAILY: {
      MONTH: 0, WEEK: 1, DATE: 2, HUB: 3,
      METRIC_INDICES: [30, 31, 32, 33, 34, 35, 36, 37, 38]
    },
    WEEKLY: {
      WEEK: 51, HUB: 52, SCORE: 77,
      METRIC_INDICES: [53, 55, 58, 61, 62, 63, 64, 65, 66],
      // Skor per metrik "LM [Backup NEW]" kolom BQ:BY (indeks 68–76), urutan sama dengan LM_METRICS
      SCORE_INDICES: [68, 69, 70, 71, 72, 73, 74, 75, 76]
    }
  },
  
  DEFAULT_ACTIONS: {
    'A1': 'Tingkatkan same day pickup ADO; tidak fm fault, pastikan pic sesuai cutoff assign, kurir attempts sesuai cutoff, dan strong onhold',
    'A2': 'Tingkatkan same day pickup seller; reminder otomatis',
    'A3': 'Audit pickup aging; follow up seller dengan pickup terlambat',
    'A4': 'Tingkatkan bulky compliance; pelatihan pengukuran',
    'A5': 'Tangani stuck parcel > 24 jam; audit inventori harian',
    'B1': 'Pastikan clean floor sebelum tengah malam; cek shift malam',
    'B2': 'Pastikan clean floor setelah tengah malam; evaluasi beban',
    'B4': 'Perbaiki SLA packed to returned; audit alur sortir',
    'C1': 'Tingkatkan daily active operator; jadwal rotasi',
    'L_A1': 'Edukasi & monitoring kurir terkait keterlambatan/pembatalan untuk menekan penalty point.',
    'L_B2': 'Optimalkan presensi & alokasi kurir 2W Dedicated & Plus harian.',
    'L_C1': 'Evaluasi rute & produktivitas bulanan kurir 2W Dedicated & Plus.',
    'L_C2': 'Evaluasi produktivitas armada 3W & 4W Dedicated.',
    'L_C3': 'Tingkatkan produktivitas operator hub sesuai target unit parcel per shift.',
    'L_D1': 'Prioritaskan pengiriman paket Sameday STD agar terkirim di hari yang sama.',
    'L_D2': 'Pantau ketat compliance pengiriman NDD ECO.',
    'L_D3': 'Percepat pengiriman paket aging D+1 Expedite Pickup.',
    'L_D4': 'Audit alasan paket return & cegah kegagalan pengiriman berulang.'
  }
};