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

// ── GET: 전체 응답 데이터 반환 ───────────────────────────────────────
function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getActiveSheet();
    const rows = sheet.getDataRange().getValues();

    if (rows.length <= 1) {
      return ContentService
        .createTextOutput(JSON.stringify({ results: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const results = rows.slice(1).map(row => {
      const answers = {};
      QUESTION_IDS.forEach((id, i) => {
        const val = row[5 + i]; // col 0~4: timestamp, name, team, jobType, tenure
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

    return ContentService
      .createTextOutput(JSON.stringify({ results }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ results: [], error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── 샘플 데이터 추가 (Apps Script 편집기에서 직접 실행) ──────────────
function loadSampleData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getActiveSheet();

  sheet.clearContents();
  sheet.appendRow(HEADERS);
  sheet.setFrozenRows(1);

  function buildAnswers(str, ppl, gov, eng, dat, cross) {
    const a = {};
    CAT_QS.STR.forEach(id   => a[id] = str);
    CAT_QS.PPL.forEach(id   => a[id] = ppl);
    CAT_QS.GOV.forEach(id   => a[id] = gov);
    CAT_QS.ENG.forEach(id   => a[id] = eng);
    CAT_QS.DAT.forEach(id   => a[id] = dat);
    CAT_QS.CROSS.forEach(id => a[id] = cross);
    return a;
  }

  const BASE_TS = new Date('2026-04-01').getTime();
  const DAY = 86400000;
  let idx = 0;

  const TENURES = ['1-3','3-5','5-7','7+','1-3','3-5','5-7','7+','1-3','3-5',
                   '3-5','5-7','1-3','7+','3-5','5-7','7+','1-3','3-5','5-7'];

  function addRow(name, team, jobType, profile, open01, open02) {
    const [str, ppl, gov, eng, dat, cross] = profile;
    const answers = buildAnswers(str, ppl, gov, eng, dat, cross);
    const ts = new Date(BASE_TS + idx * 1.5 * DAY);
    sheet.appendRow([
      Utilities.formatDate(ts, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'),
      name, team, jobType, TENURES[idx] || '3-5',
      ...QUESTION_IDS.map(id => answers[id] ?? ''),
      open01, open02
    ]);
    idx++;
  }

  const O1 = [
    'AI 도구를 업무에 적용하려 했으나 기존 ERP 시스템과 통합이 어려워 도입을 포기했습니다.',
    '팀 내에서 생성형 AI를 활용한 문서 초안 작성을 시도했으나 정보 보안 정책으로 외부 AI 서비스 사용이 금지되어 있어 실질적인 활용이 불가능했습니다.',
    'AI 분석 도구로 고객 이탈 예측 모델을 구축하려 했으나 충분한 정제 데이터가 없었고 데이터 레이블링 작업에 과도한 시간이 소요되었습니다.',
    '코파일럿 도입을 검토했으나 라이선스 비용이 예산을 초과하고 효과 측정 기준도 불명확해 경영진 승인을 받지 못했습니다.',
    '생성형 AI를 콘텐츠 제작에 사용했으나 브랜드 톤앤매너와 맞지 않아 결국 대부분을 사람이 재작성해야 했습니다.',
    'AI 챗봇을 고객 서비스에 도입했으나 한국어 특성 이해 부족으로 오답률이 높아 고객 불만이 증가했습니다.',
    '머신러닝 모델을 개발했으나 운영 환경 배포 과정에서 MLOps 인프라 부재로 모델이 실제 서비스에 적용되지 못했습니다.',
    'AI 도구 학습에 필요한 별도 교육 시간이 확보되지 않아 업무 중 틈틈이 배워야 하는 상황입니다.',
    'AI 프로젝트 ROI 측정 기준이 없어 투자 대비 성과를 경영진에게 보고하기 어렵습니다.',
    '데이터 거버넌스 체계 부재로 부서별 데이터 사일로가 심각하여 AI 모델 학습에 필요한 통합 데이터 확보가 불가능합니다.',
    'RPA 자동화 도입 후 프로세스 변경 시 재설정 비용이 과도하여 유지보수 부담이 초기 기대보다 훨씬 큽니다.',
    'AI 기반 인사 분석 도구 도입 검토 중 개인정보 보호법 준수 여부 검토에 수개월이 소요되었습니다.',
    '생산 라인 AI 품질 검사 도입 후 조명·카메라 조건에 따라 정확도가 들쑥날쑥하여 신뢰성 확보에 어려움을 겪었습니다.',
    'AI 추천 시스템 마케팅 적용 시 콜드 스타트 문제로 초기 추천 품질이 낮았습니다.',
    '자연어 처리 모델을 내부 지식 관리 시스템에 연결하려 했으나 비정형 문서 품질이 낮아 사전 정제에 예상보다 10배 이상 시간이 걸렸습니다.',
    'AI 도구 도입 시 팀원들의 디지털 기초 역량 격차가 너무 커서 기본 사용법 습득에도 어려움이 있습니다.',
    '생성형 AI 출력물의 저작권 귀속 문제와 오류 발생 시 책임 소재가 불분명하여 외부 납품물에 활용하지 못합니다.',
    '클라우드 AI 서비스 비용이 사용량에 따라 급격히 증가하여 예산 통제가 어렵습니다.',
    'AI 모델 편향성 검증 프로세스가 없어 특정 고객 집단에서 부정확한 결과가 반복적으로 발생했습니다.',
    '업무 자동화 도구 도입 후 일부 직원들이 AI로 인한 업무 대체 불안감을 갖게 되었습니다.'
  ];

  const O2 = [
    'AI 성과 측정 KPI가 없어 프로젝트 성공 여부를 판단하기 어렵습니다.',
    'AI로 인한 일자리 상실 우려가 퍼져 있어 체계적인 내부 소통과 재교육 문화 조성이 시급합니다.',
    'AI 역량 보유 직원과 미보유 직원 간 격차가 빠르게 벌어져 팀 협업 효율이 저하되고 있습니다.',
    '데이터 보안 및 개인정보 보호 기준이 AI 도입 속도를 따라가지 못하고 있습니다.',
    '실패에 대한 조직 문화적 용납이 부족해 직원들이 새로운 AI 도구 시도를 꺼립니다.',
    'AI 프로젝트가 IT 부서 주도로만 진행되어 현업 부서의 실제 문제와 괴리가 발생합니다.',
    '단기 성과 중심 평가 체계가 중장기 AI 투자를 방해합니다.',
    '중간 관리자층의 AI 이해도와 지지가 부족합니다.',
    '부서 간 데이터 공유를 가로막는 사일로 문화가 AI 모델 품질을 저해하고 있습니다.',
    'AI 윤리 기준과 책임 소재에 대한 내부 정책이 부재합니다.',
    '외부 AI 벤더 과도 의존이 내부 역량 축적을 방해합니다.',
    'AI 프로젝트 성공 사례가 조직 내에 충분히 공유되지 않아 학습이 누적되지 않습니다.',
    '현업 직원들이 자신의 업무 프로세스를 데이터화하는 방법을 모릅니다.',
    'AI 투자 예산이 특정 부서에 편중되어 있어 균형 있는 AI 역량 발전이 이루어지지 않습니다.',
    '규정 준수 부서와 AI 개발 부서 간 충돌이 잦아 도입 속도가 저하됩니다.',
    '직원들의 AI 학습 동기를 높일 인센티브 체계가 없습니다.',
    '조직 내 AI 전담 인력 부족과 높은 이직률로 인해 축적된 AI 지식이 계속 유실됩니다.',
    'AI가 창출한 가치에 대한 내부 인식 부족으로 비용 증가로 인식되는 경향이 있습니다.',
    '변화 관리 역량 부족으로 AI 도입 시 직원 저항이 반복됩니다.',
    'AI 전략과 인사 전략이 연동되지 않아 필요한 AI 인재를 적시에 확보하지 못합니다.'
  ];

  // profiles: [str, ppl, gov, eng, dat, cross]
  addRow('김민준', '재무회계', 'nondev', [3,3,4,2,3,3], O1[0],  O2[0]);
  addRow('이서연', '재무회계', 'nondev', [3,3,3,2,3,3], O1[1],  O2[1]);
  addRow('박지훈', '재무회계', 'nondev', [4,3,4,2,3,2], O1[2],  O2[2]);
  addRow('최유리', '재무회계', 'nondev', [2,3,3,2,2,3], O1[3],  O2[3]);
  addRow('정민아', '재무회계', 'nondev', [3,3,4,3,3,3], O1[4],  O2[4]);
  addRow('윤아름', '인사',     'nondev', [3,4,3,2,3,4], O1[5],  O2[5]);
  addRow('강현수', '인사',     'nondev', [3,4,3,2,2,4], O1[6],  O2[6]);
  addRow('조미래', '인사',     'nondev', [4,4,3,3,3,4], O1[7],  O2[7]);
  addRow('임성호', '인사',     'nondev', [2,3,3,2,2,3], O1[8],  O2[8]);
  addRow('한소희', '인사',     'nondev', [3,4,2,2,3,3], O1[9],  O2[9]);
  addRow('오준혁', '프론트엔드', 'dev',  [4,3,3,4,4,4], O1[10], O2[10]);
  addRow('배나연', '프론트엔드', 'dev',  [3,3,3,4,3,3], O1[11], O2[11]);
  addRow('신동현', '프론트엔드', 'dev',  [4,4,3,5,4,4], O1[12], O2[12]);
  addRow('류하은', '프론트엔드', 'dev',  [3,3,3,4,3,3], O1[13], O2[13]);
  addRow('문재원', '프론트엔드', 'dev',  [4,3,3,4,3,4], O1[14], O2[14]);
  addRow('백서진', '백엔드',   'dev',   [4,4,4,5,4,4], O1[15], O2[15]);
  addRow('허민국', '백엔드',   'dev',   [4,3,3,4,5,4], O1[16], O2[16]);
  addRow('남지혜', '백엔드',   'dev',   [4,4,4,5,5,4], O1[17], O2[17]);
  addRow('전태근', '백엔드',   'dev',   [3,3,4,4,4,3], O1[18], O2[18]);
  addRow('심유진', '백엔드',   'dev',   [3,3,3,4,4,3], O1[19], O2[19]);

  SpreadsheetApp.flush();
  Logger.log('샘플 데이터 ' + (sheet.getLastRow() - 1) + '행 추가 완료');
}
