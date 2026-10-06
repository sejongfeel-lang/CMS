// 에듀토크(담당 최세종) 겨울학기 모객 실행안 1페이지 PPT 생성 스크립트
// 실행: NODE_PATH=<pptxgenjs가 설치된 node_modules> node scripts/build_edutalk_onepager.js <출력.pptx> <apply_theme.js 경로>
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.argv[3]);
const OUT = process.argv[2] || "에듀토크_모객전략_1p_최세종.pptx";

const THEME = {
  name: "CMS Winter",
  headFontFace: "Malgun Gothic",
  bodyFontFace: "Malgun Gothic",
  colors: {
    dk1: "1B2130", lt1: "FFFFFF", dk2: "1F2A44", lt2: "EEF1F5",
    accent1: "2A7A72", accent2: "A8640F", accent3: "8A3F5C",
    accent4: "3C5A91", accent5: "5B6472", accent6: "C8553D",
    hlink: "3C5A91", folHlink: "8A3F5C",
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "에듀토크 겨울학기 모객 실행안";
const C = pres.SchemeColor;

pres.defineSlideMaster({
  title: "ONE_PAGER",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.4, y: 0.28, w: 9.2, h: 0.5,
      fontSize: 22, bold: true, color: C.text2, margin: 0, align: "left", valign: "middle" }, text: "" } },
  ],
});

pres.addSection({ title: "에듀토크" });
const s = pres.addSlide({ masterName: "ONE_PAGER", sectionTitle: "에듀토크" });
const T = (text, opts) => s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontSize: 9, color: C.text1 }, opts));
const tint = (x, y, w, h, color, name) => s.addShape(pres.shapes.RECTANGLE, { x, y, w, h,
  fill: { color, transparency: 92 }, line: { type: "none" }, objectName: name });
const head = (x, y, w, color, label, sub) => {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.34, fill: { color }, line: { type: "none" }, rectRadius: 0.06 });
  T([{ text: label, options: { bold: true, fontSize: 10.5 } }].concat(sub ? [{ text: "  " + sub, options: { fontSize: 8.5 } }] : []),
    { x: x + 0.1, y, w: w - 0.2, h: 0.34, color: C.background1, valign: "middle" });
};

s.addText([{ text: "에듀토크 (이병훈TV 라이브)" }, { text: "   겨울학기 모객 실행 방향 · 담당 최세종", options: { fontSize: 13, bold: false } }], { placeholder: "title" });
T("10/7(수) 대회의실 · 방송 D-Day 11/4(수)", { x: 9.7, y: 0.28, w: 3.23, h: 0.5, align: "right", valign: "middle", color: C.accent5 });

// ---------- WHY / WHO / HOW ----------
const LX = 0.4, LW = 8.35, GAP = 0.12, CW = (LW - GAP * 2) / 3, R1 = 0.95, R1H = 1.42;
const cards = [
  [C.accent1, "왜 하나", "겨울학기 모객의 첫 대규모 학부모 접점. 이병훈TV의 도달력·신뢰도로 입시 관심 학부모 DB를 온라인에서 한 번에 확보하고, 인뎁스 세미나·체험수업·입학테스트로 이어지는 퍼널의 입구를 만든다."],
  [C.accent2, "누구에게", "입시·진학 정보에 관심 있는 학부모 (세부 타깃은 주제·패널 확정 후 확정, 홍혜숙·석다혜 10/6). 이병훈TV 시청자 + 제휴학원(기존 입시 실적) + 광고 채널 유입."],
  [C.accent3, "어떻게", "11/4 이병훈TV 라이브 방송 · 랜딩 웹앱에서 신청(설문 + 특전) · DM 2회(D-14 예열, D-7 리마인드, 알림톡 분할 발송) · 방송 중 라이브 채팅 실시간 대응."],
];
cards.forEach(([col, h, body], i) => {
  const x = LX + i * (CW + GAP);
  tint(x, R1, CW, R1H, col, "card-" + h);
  head(x, R1, CW, col, h);
  T(body, { x: x + 0.12, y: R1 + 0.42, w: CW - 0.24, h: R1H - 0.48, fontSize: 9, valign: "top", lineSpacingMultiple: 1.05 });
});

// ---------- 모객 연결 흐름 ----------
const R2 = 2.52;
T("모객 연결 흐름", { x: LX, y: R2, w: 3, h: 0.26, fontSize: 10.5, bold: true, color: C.text2, valign: "middle" });
const steps = [
  ["홍보 채널·제휴", "10/5~ 채널 확정\n제휴학원 공동 홍보"],
  ["랜딩 신청", "10/21 오픈\n설문+특전 → DB"],
  ["DM 예열", "D-14 (10/21)\nD-7 (10/28)"],
  ["11/4 라이브", "이병훈TV 방송\n채팅 실시간 대응"],
  ["후속 알림톡", "방송 후 인뎁스\n·체험수업 안내"],
  ["입테·등록", "12월 2주 입테\n→ 겨울학기 등록"],
];
const AG = 0.24, SW = (LW - AG * (steps.length - 1)) / steps.length, SH = 0.86, SY = R2 + 0.32;
const stepColors = [C.accent1, C.accent1, C.accent2, C.accent2, C.accent3, C.accent3];
steps.forEach(([h, d], i) => {
  const x = LX + i * (SW + AG);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: SY, w: SW, h: SH, rectRadius: 0.06,
    fill: { color: stepColors[i] }, line: { type: "none" }, objectName: "step-" + h });
  T([{ text: h, options: { bold: true, fontSize: 10, breakLine: true } }, { text: d, options: { fontSize: 8 } }],
    { x: x + 0.1, y: SY, w: SW - 0.2, h: SH, color: C.background1, valign: "middle" });
  if (i < steps.length - 1)
    T("▶", { x: x + SW, y: SY, w: AG, h: SH, fontSize: 10, color: C.accent5, align: "center", valign: "middle" });
});

// ---------- 주요 일정 (최세종 담당) ----------
const R3 = 3.88, R3H = 2.42, SCW = 4.25;
head(LX, R3, SCW, C.text2, "주요 일정", "최세종 담당 업무");
const sched = [
  ["10/5", "홍보 채널 확정 (제휴실행안·제휴학원 체크)", "진행중"],
  ["10/8", "특전 설정 (주제 확정 후 컨셉 결정)", ""],
  ["10/13", "랜딩 페이지 웹앱 구성 기획 (+최유나)", ""],
  ["10/14", "라이브 채팅 대응방안 기획", ""],
  ["10/16", "설문 설계 · DM 발송 문안 개발", ""],
  ["10/19", "랜딩 페이지 웹앱 개발", ""],
  ["10/21", "랜딩 릴리즈 · DM 1차 발송 (D-14)", ""],
  ["10/28", "DM 2차 리마인드 (D-7)", ""],
  ["11/4", "방송 D-Day · 라이브 채팅 대응", ""],
];
const rows = sched.map(([d, t, st], i) => [
  { text: d, options: { bold: true, color: C.text2 } },
  { text: t },
  { text: st, options: { color: C.accent1, bold: true, align: "center" } },
]);
s.addTable(rows, { x: LX, y: R3 + 0.4, w: SCW, colW: [0.55, 3.05, 0.65], rowH: 0.218, fontSize: 8.5,
  fontFace: THEME.bodyFontFace, color: C.text1, valign: "middle", margin: [0, 0.05, 0, 0.05],
  border: { type: "solid", pt: 0.5, color: "D4D8DE" }, fill: { color: C.background1 } });

// ---------- 예상 예산 ----------
const BX = LX + SCW + 0.15, BW = LW - SCW - 0.15;
head(BX, R3, BW, C.accent6, "예상 예산", "금액은 회의 전 기입");
const budget = [
  ["항목", "내용", "금액(만원)"],
  ["특전 개발·제작", "주제 연계 특전 (최유나 10/16)", ""],
  ["홍보 채널 광고", "채널별 광고 소재·집행", ""],
  ["제휴 홍보", "제휴학원 공동 홍보 (제휴실행안)", ""],
  ["DM·알림톡 발송", "D-14 / D-7, 분할 발송", ""],
  ["방송·패널", "이병훈TV 송출·패널 (조율 중)", ""],
  ["랜딩 웹앱", "내부 개발 (최세종)", "-"],
  ["합계", "", ""],
];
const brows = budget.map((r, i) => r.map((v, j) => ({ text: v, options: {
  bold: i === 0 || i === budget.length - 1 || j === 0,
  fill: { color: i === 0 ? "F2E3DF" : "FFFFFF" }, align: j === 2 ? "center" : "left" } })));
s.addTable(brows, { x: BX, y: R3 + 0.4, w: BW, colW: [1.05, 2.0, BW - 3.05], rowH: 0.245, fontSize: 8.5,
  fontFace: THEME.bodyFontFace, color: C.text1, valign: "middle", margin: [0, 0.05, 0, 0.05],
  border: { type: "solid", pt: 0.5, color: "D4D8DE" } });

// ---------- 추가 제안 ----------
const PX = 8.92, PW = 4.01, PY = 0.95, PH = 5.35;
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: PX, y: PY, w: PW, h: PH,
  fill: { color: C.text2 }, line: { type: "none" }, rectRadius: 0.06, objectName: "proposal-panel" });
T([
  { text: "✚ 추가 제안", options: { bold: true, fontSize: 11 } },
  { text: "  12월 모객 로드맵 적용", options: { fontSize: 8.5, color: "CADCFC" } },
], { x: PX + 0.15, y: PY, w: PW - 0.3, h: 0.4, color: C.background1, valign: "middle" });
const props = [
  ["A. 리마인더 확대로 미시청 방지", "D-14·D-7에 D-1·당일 알림톡 추가 → 신청 후 미시청(노쇼) 최소화"],
  ["B. 설문 기반 우선순위 상담콜", "설문 관심도로 신청자 스코어링 → 방송 후 24~48h 골든타임에 상위 리드부터 상담콜"],
  ["C. 관심사별 후속 알림톡 분기", "설문의 학년·관심 프로그램에 따라 인뎁스 세미나 / FIT / MATHMILE 체험수업으로 나눠 안내"],
  ["D. 다시보기·하이라이트 숏폼", "미시청자 재접촉용 다시보기 링크 + 하이라이트를 메타·당근 광고 소재로 재활용"],
  ["E. 입테 통합 퍼널·CRM 연결", "에듀토크 DB를 신청→시청→인뎁스·체험→입테→등록까지 추적해 단계별 전환율 측정"],
  ["F. 라이브 채팅 질문 FAQ화", "채팅 질문을 학부모 FAQ·상담 스크립트·후속 콘텐츠로 정리 (Montelena 여론 데이터와 결합)"],
];
const pruns = [];
props.forEach(([h, d], i) => {
  pruns.push({ text: h, options: { bold: true, fontSize: 9.5, color: "FFFFFF", paraSpaceBefore: i ? 9 : 0, breakLine: true } });
  pruns.push({ text: d, options: { fontSize: 8.5, color: "DCE3EE", breakLine: i < props.length - 1 } });
});
T(pruns, { x: PX + 0.15, y: PY + 0.5, w: PW - 0.3, h: PH - 0.6, valign: "top", lineSpacingMultiple: 1.05 });

// ---------- 회의 결정 필요 ----------
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.4, y: 6.44, w: 12.53, h: 0.6,
  fill: { color: C.accent6, transparency: 88 }, line: { color: C.accent6, width: 0.75 }, rectRadius: 0.06, objectName: "decisions" });
T("회의 결정 필요", { x: 0.55, y: 6.44, w: 1.3, h: 0.6, fontSize: 10.5, bold: true, color: C.accent6, valign: "middle" });
T("① 이병훈TV 컨셉 대표님 보고·조율 결과   ② 주제·패널 확정 → 특전 컨셉(10/8)   ③ 홍보 채널·제휴학원 범위   ④ 당근마켓 홍보 품의 에듀토크 통합 여부   ⑤ 알림톡 분할 발송 기준·템플릿 심사 일정   ⑥ 성과 목표(신청·시청·입테 전환)",
  { x: 1.85, y: 6.44, w: 10.95, h: 0.6, fontSize: 9, valign: "middle" });
T("출처: CMS MKT | 2026 4Q · [4Q] Items통합 B~H열 에듀토크(10/6 기준) / 12월 모객 로드맵(9/9, Claude Code 작업) · D-7 리마인드일(10/28)은 방송일 기준 산출",
  { x: 0.4, y: 7.1, w: 12.53, h: 0.22, fontSize: 7.5, color: C.accent5 });

s.addNotes("[4Q] Items통합에서 담당자가 최세종인 아이템(에듀토크)만 정리. 협업 업무(주제 기획·패널: 홍혜숙/석다혜, 광고 소재·특전 제작: 최유나, 비주얼: 공혜영)는 일정 의존 관계로만 언급. 예산 금액은 시트에 없어 회의 전 기입 필요.");

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
