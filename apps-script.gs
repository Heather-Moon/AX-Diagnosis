// AX 조직역량 진단 — Google Sheets 연동 스크립트
// 사용법: Google Sheets > 확장 프로그램 > Apps Script에 붙여넣고 웹 앱으로 배포

const SHEET_ID = '14HfbcXcdGMg_eKVn6_nb0KH8cv2KCm0ldkci5TP_RR8';

const QUESTION_IDS = [
  'STR-P-01', 'STR-T-01', 'STR-O-01', 'STR-L-01', 'STR-O-02',
  'PPL-O-01', 'PPL-O-02', 'PPL-T-01', 'PPL-T-02', 'PPL-P-01', 'PPL-P-02', 'PPL-L-01', 'PPL-L-02',
  'GOV-O-01', 'GOV-O-02', 'GOV-P-01', 'GOV-T-01', 'GOV-O-03',
  'ENG-P-01', 'ENG-T-01', 'ENG-O-01', 'ENG-O-02', 'ENG-P-02',
  'DAT-P-01', 'DAT-O-01', 'DAT-O-02', 'DAT-T-01', 'DAT-O-03',
  'CROSS-L-01', 'CROSS-T-01', 'CROSS-P-01', 'CROSS-T-02'
];

const CAT_QS = {
  STR:   ['STR-P-01','STR-T-01','STR-O-01','STR-L-01','STR-O-02'],
  PPL:   ['PPL-O-01','PPL-O-02','PPL-T-01','PPL-T-02','PPL-P-01','PPL-P-02','PPL-L-01','PPL-L-02'],
  GOV:   ['GOV-O-01','GOV-O-02','GOV-P-01','GOV-T-01','GOV-O-03'],
  ENG:   ['ENG-P-01','ENG-T-01','ENG-O-01','ENG-O-02','ENG-P-02'],
  DAT:   ['DAT-P-01','DAT-O-01','DAT-O-02','DAT-T-01','DAT-O-03'],
  CROSS: ['CROSS-L-01','CROSS-T-01','CROSS-P-01','CROSS-T-02'],
};

const OPEN_IDS = ['OPEN-01', 'OPEN-02'];

const HEADERS = [
  '타임스탬프', '이름', '팀', '직군', '연차',
  ...QUESTION_IDS,
  ...OPEN_IDS
];

// ── POST: 설문 응답 저장 ─────────────────────────────────────────────
function doPost(e) {
  try {
    let data;
    if (e.postData && e.postData.contents && e.postData.contents.trim().startsWith('{')) {
      data = JSON.parse(e.postData.contents);
    } else if (e.parameter && e.parameter.payload) {
      data = JSON.parse(e.parameter.payload);
    } else {
      throw new Error('No valid data received');
    }
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getActiveSheet();

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
    }

    const ts = new Date(data.timestamp);
    const row = [
      Utilities.formatDate(ts, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'),
      data.name || '익명',
      data.team || '',
      data.jobType || '',
      data.tenure || '',
      ...QUESTION_IDS.map(id => (data.answers || {})[id] ?? ''),
      ...OPEN_IDS.map(id => (data.openAnswers || {})[id] ?? '')
    ];

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── GET: 전체 응답 데이터 반환 (JSONP 지원) ─────────────────────────
function doGet(e) {
  let out;
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getActiveSheet();
    const rows = sheet.getDataRange().getValues();

    if (rows.length <= 1) {
      out = JSON.stringify({ results: [] });
    } else {
      const results = rows.slice(1).map(row => {
        const answers = {};
        QUESTION_IDS.forEach((id, i) => {
          const val = row[5 + i];
          answers[id] = typeof val === 'number' ? val : (parseInt(val) || 0);
        });
        const openAnswers = {};
        OPEN_IDS.forEach((id, i) => {
          openAnswers[id] = String(row[5 + QUESTION_IDS.length + i] || '');
        });
        const tsCell = row[0];
        const ts = tsCell instanceof Date ? tsCell.getTime() : new Date(tsCell).getTime();
        return {
          timestamp: isNaN(ts) ? Date.now() : ts,
          name: String(row[1] || '익명'),
          team: String(row[2] || ''),
          jobType: String(row[3] || 'nondev'),
          tenure: String(row[4] || ''),
          answers,
          openAnswers
        };
      });
      out = JSON.stringify({ results });
    }
  } catch (err) {
    out = JSON.stringify({ results: [], error: err.toString() });
  }

  const cb = e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService
      .createTextOutput(cb + '(' + out + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(out)
    .setMimeType(ContentService.MimeType.JSON);
}
