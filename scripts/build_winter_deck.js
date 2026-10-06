// 겨울학기 모객 회의 자료 PPT 생성 스크립트 (6장)
//  1. 12월 모객 아이템 현황 (업무 진행 상태·주차별 업무량 차트)
//  2. CMS 프로젝트 허브 구축 현황 (아이템별 운영 기간 타임라인·입력 현황)
//  3. 에듀토크 실행 방향 (담당 최세종)
//  4~6. 에듀토크 상세: 준비 일정·크리티컬 패스 / 방송 운영안 / 모객 시퀀스·전환 설계
// 실행: NODE_PATH=<pptxgenjs가 설치된 node_modules> node scripts/build_winter_deck.js <출력.pptx> <apply_theme.js 경로>
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.argv[3]);
const OUT = process.argv[2] || "겨울학기_모객_회의자료.pptx";

const THEME = {
  name: "CMS Winter",
  headFontFace: "Pretendard",
  bodyFontFace: "Pretendard",
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
pres.title = "겨울학기 모객 회의 자료";
const C = pres.SchemeColor;

pres.defineSlideMaster({
  title: "ONE_PAGER",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.4, y: 0.28, w: 9.2, h: 0.5,
      fontSize: 22, bold: true, color: C.text2, margin: 0, align: "left", valign: "middle" }, text: "" } },
  ],
});

let s;
const T = (text, opts) => s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontSize: 9, color: C.text1 }, opts));
const tint = (x, y, w, h, color, name) => s.addShape(pres.shapes.RECTANGLE, { x, y, w, h,
  fill: { color, transparency: 92 }, line: { type: "none" }, objectName: name });
const head = (x, y, w, color, label, sub) => {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.34, fill: { color }, line: { type: "none" }, rectRadius: 0.06 });
  T([{ text: label, options: { bold: true, fontSize: 10.5 } }].concat(sub ? [{ text: "  " + sub, options: { fontSize: 8.5 } }] : []),
    { x: x + 0.1, y, w: w - 0.2, h: 0.34, color: C.background1, valign: "middle" });
};
const footer = (text) => T(text, { x: 0.4, y: 7.1, w: 12.53, h: 0.22, fontSize: 7.5, color: C.accent5 });
const HEX = THEME.colors;
const CHART_TXT = { catAxisLabelColor: "1B2130", valAxisLabelColor: "5B6472", catAxisLabelFontFace: "+mn-lt",
  valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt", titleFontFace: "+mn-lt",
  catAxisLabelFontSize: 9, valAxisLabelFontSize: 8, dataLabelFontSize: 8, legendFontSize: 9,
  valGridLine: { color: "E3E6EB", size: 0.5 }, catGridLine: { style: "none" } };

// 단계 구분 (허브 15개 아이템 + 통합 아웃바운드)
const STAGE = { awr: [C.accent1, "인지·DB 확보"], eng: [C.accent2, "체험·관계 형성"], cnv: [C.accent3, "응시·등록 전환"], out: [C.accent4, "전환 지원"] };
// [허브 아이템명, 시트 PM, 단계, 시작, 종료, 업무수, 완료, 진행중, 대기]  — [4Q] Items통합 B~H열(10/6) 집계
const ITEMS = [
  ["에듀토크", "최세종", "awr", "09-22", "11-04", 24, 1, 4, 0],
  ["메타 광고 (인스타그램)", "최유나", "awr", "10-06", "10-26", 14, 0, 0, 0],
  ["인플루언서 마케팅", "최유나", "awr", "09-23", "12-11", 20, 1, 2, 0],
  ["당근 마켓", "석다혜", "awr", "10-12", "11-13", 8, 0, 0, 0],
  ["네이버 카페", "이미현", "awr", "10-02", "11-20", 9, 0, 0, 0],
  ["KMO 2차 현장 응원", "김새미", "awr", "10-08", "11-14", 8, 0, 0, 1],
  ["하반기 자사고 입학설명회 현장", "PM 미기재", "awr", "10-08", "10-08", 18, 14, 2, 0],
  ["인뎁스 세미나", "석다혜", "eng", "10-08", "11-03", 9, 0, 0, 0],
  ["아샘", "이미현", "eng", "10-06", "10-28", 5, 0, 0, 0],
  ["체험수업 (FIT·MATHMILE)", "석다혜/이미현", "eng", "10-02", "10-30", 13, 0, 1, 0],
  ["자사고 면접 컨설팅", "홍혜숙", "eng", null, null, 2, 0, 2, 0],
  ["교과 프로그램 (알림톡)", "이미현/석다혜", "cnv", "10-19", "11-03", 9, 0, 0, 0],
  ["친구 추천 이벤트", "이로은", "cnv", "10-15", "12-11", 18, 0, 2, 0],
  ["수강후기 프로모션", "최유나", "cnv", "10-07", "11-19", 16, 1, 0, 0],
  ["현장 주도 마케팅", "최유나", "cnv", "09-29", "12-18", 6, 1, 0, 0],
];

// ================= 1. 12월 모객 아이템 현황 =================
pres.addSection({ title: "현황" });
s = pres.addSlide({ masterName: "ONE_PAGER", sectionTitle: "현황" });
s.addText([{ text: "12월 모객 아이템 현황" }, { text: "   15개 아이템 · 179개 업무 한눈에 보기", options: { fontSize: 13, bold: false } }], { placeholder: "title" });
T("10/7(수) 11:00 대회의실 · 마케팅팀", { x: 9.7, y: 0.28, w: 3.23, h: 0.5, align: "right", valign: "middle", color: C.accent5 });

const KPI = [
  ["15 + 1", "확정 아이템 + 통합 아웃바운드", "CMS 프로젝트 허브 기준", C.text2],
  ["179건", "세부 업무", "[4Q] Items통합 기준", C.accent4],
  ["18 · 13", "완료 · 진행중", "148건 착수 전", C.accent1],
  ["68%", "3주(10/5~25)에 마감 집중", "136건 중 93건", C.accent6],
];
const KW = (12.53 - 0.2 * 3) / 4;
KPI.forEach(([big, lbl, sub, col], i) => {
  const x = 0.4 + i * (KW + 0.2);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 0.98, w: KW, h: 1.05, rectRadius: 0.06,
    fill: { color: C.background2 }, line: { type: "none" }, objectName: "kpi-" + lbl });
  T(big, { x: x + 0.2, y: 1.05, w: 1.4, h: 0.9, fontSize: 26, bold: true, color: col, valign: "middle" });
  T([{ text: lbl, options: { bold: true, fontSize: 10, breakLine: true } }, { text: sub, options: { fontSize: 8.5, color: C.accent5 } }],
    { x: x + 1.65, y: 1.05, w: KW - 1.75, h: 0.9, valign: "middle" });
});

// 아이템별 업무 진행 상태 (누적 가로 막대)
const byTotal = ITEMS.slice().sort((a, b) => b[5] - a[5]);
const short = (n) => n.replace(" (인스타그램)", "").replace(" (FIT·MATHMILE)", "").replace(" (알림톡)", "").replace("하반기 자사고 입학설명회 현장", "자사고 설명회 현장");
const labels = byTotal.map((r) => short(r[0]));
T("아이템별 업무 진행 상태", { x: 0.4, y: 2.25, w: 6.9, h: 0.3, fontSize: 11, bold: true, color: C.text2 });
s.addChart(pres.charts.BAR, [
  { name: "완료", labels, values: byTotal.map((r) => r[6]) },
  { name: "진행중", labels, values: byTotal.map((r) => r[7]) },
  { name: "대기", labels, values: byTotal.map((r) => r[8]) },
  { name: "착수 전", labels, values: byTotal.map((r) => r[5] - r[6] - r[7] - r[8]) },
], Object.assign({}, CHART_TXT, {
  x: 0.4, y: 2.55, w: 6.9, h: 3.85, barDir: "bar", barGrouping: "stacked", barGapWidthPct: 45,
  catAxisOrientation: "maxMin", valAxisHidden: true, valGridLine: { style: "none" },
  chartColors: [HEX.accent1, HEX.accent2, HEX.accent6, "C9CED6"],
  showLegend: true, legendPos: "t", showValue: true, dataLabelPosition: "ctr", dataLabelColor: "FFFFFF",
  dataLabelFormatCode: "#,##0;;;", objectName: "chart-status",
}));

// 주차별 업무 마감 건수 (세로 막대)
const WEEKS = [["9/21", 2], ["9/28", 3], ["10/5", 25], ["10/12", 39], ["10/19", 29], ["10/26", 13], ["11/2", 9], ["11/9", 6], ["11/16", 2], ["11/23", 0], ["11/30", 2], ["12/7", 5], ["12/14", 1]];
T("주차별 업무 마감 건수", { x: 7.6, y: 2.25, w: 5.33, h: 0.3, fontSize: 11, bold: true, color: C.text2 });
s.addChart(pres.charts.BAR, [{ name: "마감 업무", labels: WEEKS.map((w) => w[0]), values: WEEKS.map((w) => w[1]) }],
  Object.assign({}, CHART_TXT, {
    x: 7.6, y: 2.55, w: 5.33, h: 2.75, barDir: "col", barGapWidthPct: 35, catAxisLabelFontSize: 8,
    chartColors: [HEX.accent4], showLegend: false, showValue: true, dataLabelPosition: "outEnd",
    dataLabelColor: "1B2130", valAxisHidden: true, valGridLine: { style: "none" }, objectName: "chart-weekly",
  }));
T("주 시작일(월) 기준 · 일정이 기재된 136건", { x: 7.6, y: 5.3, w: 5.33, h: 0.2, fontSize: 8, color: C.accent5 });
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.6, y: 5.6, w: 5.33, h: 0.8, rectRadius: 0.06,
  fill: { color: C.accent6, transparency: 88 }, line: { type: "none" }, objectName: "weekly-note" });
T([
  { text: "10/12 주 39건 피크", options: { bold: true, color: C.accent6, breakLine: true } },
  { text: "에듀토크·메타 광고·체험수업·친구추천 준비가 한 주에 겹침 → 디자인 요청과 품의 일정 우선순위 조정 필요" },
], { x: 7.75, y: 5.6, w: 5.03, h: 0.8, fontSize: 9, valign: "middle" });

s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.4, y: 6.55, w: 12.53, h: 0.48,
  fill: { color: C.background2 }, line: { type: "none" }, rectRadius: 0.06, objectName: "takeaway" });
T([
  { text: "시사점  ", options: { bold: true, color: C.text2 } },
  { text: "아이템 대부분이 착수 전 → 오늘 실행 방향 확정이 우선 · 업무 담당 상위: 최유나 39건, 석다혜 38건, 이미현 26건 → 10월 중순 업무 분산 검토" },
], { x: 0.55, y: 6.55, w: 12.25, h: 0.48, fontSize: 9.5, valign: "middle" });
footer("출처: CMS MKT | 2026 4Q · [4Q] Items통합 B~H열(10/6 기준) 업무 행 집계 · '착수 전'은 완료여부 칸이 비어 있는 업무");

// ================= 2. CMS 프로젝트 허브 구축 현황 =================
pres.addSection({ title: "허브 구축" });
s = pres.addSlide({ masterName: "ONE_PAGER", sectionTitle: "허브 구축" });
s.addText([{ text: "CMS 프로젝트 허브" }, { text: "   12월 모객 아이템(HQ) 구축 현황", options: { fontSize: 13, bold: false } }], { placeholder: "title" });
T("허브 화면 기준 · 시트 일정 대조", { x: 9.7, y: 0.28, w: 3.23, h: 0.5, align: "right", valign: "middle", color: C.accent5 });

// 운영 기간 타임라인
const GX = 0.4, GW = 8.35, LBL = 2.75, AX = GX + LBL, AW = GW - LBL - 0.1;
const D0 = new Date(2026, 8, 21), D1 = new Date(2026, 11, 21);
const xOf = (md) => { const [m, d] = md.split("-").map(Number); return AX + ((new Date(2026, m - 1, d) - D0) / (D1 - D0)) * AW; };
T("아이템별 운영 기간 (허브 ‘시기’ 제안)", { x: GX, y: 0.92, w: 3.5, h: 0.3, fontSize: 11, bold: true, color: C.text2 });
const legend = [STAGE.awr, STAGE.eng, STAGE.cnv, STAGE.out];
legend.forEach(([col, name], i) => {
  const lx = GX + 3.65 + i * 1.18;
  s.addShape(pres.shapes.OVAL, { x: lx, y: 1.01, w: 0.12, h: 0.12, fill: { color: col }, line: { type: "none" } });
  T(name, { x: lx + 0.17, y: 0.94, w: 1.1, h: 0.26, fontSize: 8.5, valign: "middle" });
});
const GY = 1.62, RH = 0.29, rowsAll = ITEMS.concat([["통합 아웃바운드", "HQ", "out", null, null, 0, 0, 0, 0]]);
[["10-01", "10월"], ["11-01", "11월"], ["12-01", "12월"]].forEach(([md, lbl]) => {
  const x = xOf(md);
  s.addShape(pres.shapes.LINE, { x, y: GY - 0.05, w: 0, h: RH * rowsAll.length + 0.05, line: { color: "C9CED6", width: 0.75, dashType: "dash" } });
  T(lbl, { x: x + 0.04, y: GY - 0.3, w: 0.5, h: 0.22, fontSize: 8.5, bold: true, color: C.accent5 });
});
const todayX = xOf("10-07");
s.addShape(pres.shapes.LINE, { x: todayX, y: GY - 0.05, w: 0, h: RH * rowsAll.length + 0.05, line: { color: HEX.accent6, width: 1.25 } });
T("오늘 10/7", { x: todayX - 0.45, y: GY + RH * rowsAll.length + 0.02, w: 0.9, h: 0.2, fontSize: 8, bold: true, color: C.accent6, align: "center" });
rowsAll.forEach(([name, pm, st, a, b, n], i) => {
  const y = GY + i * RH, col = STAGE[st][0];
  if (i % 2 === 0) s.addShape(pres.shapes.RECTANGLE, { x: GX, y, w: GW, h: RH, fill: { color: C.background2, transparency: 40 }, line: { type: "none" } });
  s.addShape(pres.shapes.OVAL, { x: GX + 0.06, y: y + RH / 2 - 0.05, w: 0.1, h: 0.1, fill: { color: col }, line: { type: "none" } });
  T([{ text: name, options: { bold: true, fontSize: 8.5 } }, { text: "  " + pm, options: { fontSize: 7.5, color: C.accent5 } }],
    { x: GX + 0.22, y, w: LBL - 0.25, h: RH, valign: "middle" });
  if (a) {
    const x1 = xOf(a), x2 = Math.max(xOf(b), x1 + 0.08);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x1, y: y + 0.07, w: x2 - x1, h: RH - 0.14, rectRadius: 0.04,
      fill: { color: col }, line: { type: "none" }, objectName: "period-" + name });
    const lbl = (a === b ? a : a + " ~ " + b).replace(/(\d+)-(\d+)/g, (m, mm, dd) => +mm + "/" + +dd) + (n ? " · " + n + "건" : "");
    if (x2 + 1.5 > GX + GW) T(lbl, { x: x2 - 1.95, y, w: 1.88, h: RH, fontSize: 7.5, bold: true, color: C.background1, align: "right", valign: "middle" });
    else T(lbl, { x: x2 + 0.06, y, w: 1.9, h: RH, fontSize: 7.5, color: C.accent5, valign: "middle" });
  } else {
    T(st === "out" ? "시기 미정 · 추가 제안 C(콜 스코어링)와 연결" : "별도 마스터시트로 일정 관리", { x: AX + 0.05, y, w: 4, h: RH, fontSize: 7.5, color: C.accent5, valign: "middle", italic: true });
  }
});

// 허브 입력 현황 패널
const HX = 8.95, HW = 3.98, HY = 0.92, HH = 5.5;
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: HX, y: HY, w: HW, h: HH, rectRadius: 0.06,
  fill: { color: C.text2 }, line: { type: "none" }, objectName: "hub-panel" });
T("허브 입력 현황", { x: HX + 0.2, y: HY + 0.05, w: HW - 0.4, h: 0.4, fontSize: 11, bold: true, color: C.background1, valign: "middle" });
const meters = [["아이템 등록", 16, 16, "아이템 14 · WEB 2"], ["시기 입력", 0, 16, "전 항목 ‘시기 미지정’"], ["센터 담당 배정 (정·부·실)", 0, 16, "항목별 미선택 38"]];
meters.forEach(([lbl, v, tot, note], i) => {
  const y = HY + 0.6 + i * 0.68, bw = HW - 0.4;
  T([{ text: lbl, options: { bold: true } }, { text: "   " + v + " / " + tot, options: { color: "CADCFC" } }],
    { x: HX + 0.2, y, w: bw, h: 0.22, fontSize: 9, color: C.background1 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: HX + 0.2, y: y + 0.27, w: bw, h: 0.12, rectRadius: 0.06, fill: { color: "3A4766" }, line: { type: "none" } });
  if (v) s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: HX + 0.2, y: y + 0.27, w: bw * v / tot, h: 0.12, rectRadius: 0.06, fill: { color: "6FBF87" }, line: { type: "none" } });
  T(note, { x: HX + 0.2, y: y + 0.42, w: bw, h: 0.2, fontSize: 8, color: "DCE3EE" });
});
T("다음 단계", { x: HX + 0.2, y: HY + 2.75, w: HW - 0.4, h: 0.28, fontSize: 10, bold: true, color: C.background1 });
const nexts = [
  "왼쪽 운영 기간을 허브 ‘시기’에 입력",
  "아이템별 센터 담당 정·부·실 배정",
  "통합 아웃바운드(WEB)에 미전환 DB 콜 흐름 연결 — 우선순위 산출 → 알림톡 예열 → 상담콜",
  "명칭 통일: 허브 ‘교과 프로그램’ = 시트 ‘교과 프로그램 홍보(알림톡)’, 허브 ‘자사고 면접 컨설팅(이병훈 소장 라이브)’ = 시트 ‘[특사모]’ + ‘이병훈 소장 라이브’",
];
T(nexts.map((t, i) => ({ text: t, options: { bullet: { type: "number" }, breakLine: i < nexts.length - 1, paraSpaceAfter: 5 } })),
  { x: HX + 0.2, y: HY + 3.05, w: HW - 0.4, h: 2.35, fontSize: 8.5, color: "DCE3EE", valign: "top" });

s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.4, y: 6.55, w: 12.53, h: 0.48,
  fill: { color: C.background2 }, line: { type: "none" }, rectRadius: 0.06, objectName: "takeaway" });
T([
  { text: "시사점  ", options: { bold: true, color: C.text2 } },
  { text: "허브에 16개 항목 등록은 끝났고, 시기·센터 담당이 비어 있음 → 시트 일정 기준으로 이번 주 안에 채우면 센터별 실행 현황까지 허브에서 추적 가능" },
], { x: 0.55, y: 6.55, w: 12.25, h: 0.48, fontSize: 9.5, valign: "middle" });
footer("출처: CMS 프로젝트 허브 ‘12월 모객 아이템(HQ)’·‘아웃바운드(HQ)’ 화면 / [4Q] Items통합 B~H열 업무 날짜의 최초~최종일(10/6 기준)");

// ================= 3. 에듀토크 =================
pres.addSection({ title: "에듀토크" });
s = pres.addSlide({ masterName: "ONE_PAGER", sectionTitle: "에듀토크" });
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
head(LX, R3, SCW, C.text2, "주요 일정", "최세종 담당 업무 타임라인");
const sched = [
  ["10/5", "홍보 채널 확정", "진행중"],
  ["10/8", "특전 설정", ""],
  ["10/13", "랜딩 구성 기획", ""],
  ["10/14", "채팅 대응방안 기획", ""],
  ["10/16", "설문 설계·DM 문안", ""],
  ["10/19", "랜딩 웹앱 개발", ""],
  ["10/21", "랜딩 릴리즈·DM 1차", ""],
  ["10/28", "DM 2차 (D-7)", ""],
  ["11/4", "방송 D-Day·채팅 대응", ""],
];
const MX = LX + 2.35, MW = SCW - 2.5, MD0 = new Date(2026, 9, 1), MD1 = new Date(2026, 10, 7);
const mx = (md) => { const [m, d] = md.split("/").map(Number); return MX + ((new Date(2026, m - 1, d) - MD0) / (MD1 - MD0)) * MW; };
const MY = R3 + 0.62, MRH = 0.195;
[["10/1", "10/1"], ["10/15", "10/15"], ["11/1", "11/1"]].forEach(([md, lbl]) => {
  T(lbl, { x: mx(md) - 0.25, y: R3 + 0.4, w: 0.5, h: 0.18, fontSize: 7.5, color: C.accent5, align: "center" });
  s.addShape(pres.shapes.LINE, { x: mx(md), y: MY - 0.03, w: 0, h: MRH * sched.length, line: { color: "C9CED6", width: 0.5, dashType: "dash" } });
});
s.addShape(pres.shapes.LINE, { x: mx("10/7"), y: MY - 0.03, w: 0, h: MRH * sched.length, line: { color: HEX.accent6, width: 1 } });
sched.forEach(([d, t, st], i) => {
  const y = MY + i * MRH, done = !!st, last = i === sched.length - 1;
  T([{ text: d + "  ", options: { bold: true, color: C.text2 } }, { text: t }].concat(st ? [{ text: "  " + st, options: { bold: true, color: C.accent1 } }] : []),
    { x: LX, y, w: 2.35, h: MRH, fontSize: 8, valign: "middle" });
  s.addShape(pres.shapes.LINE, { x: MX, y: y + MRH / 2, w: MW, h: 0, line: { color: "E3E6EB", width: 0.5 } });
  s.addShape(pres.shapes.OVAL, { x: mx(d) - 0.055, y: y + MRH / 2 - 0.055, w: 0.11, h: 0.11,
    fill: { color: last ? HEX.accent3 : done ? HEX.accent1 : "FFFFFF" }, line: { color: last ? HEX.accent3 : HEX.dk2, width: 1 } });
});
T("● 진행중   ○ 예정   │ 빨간선 오늘(10/7)", { x: LX, y: MY + MRH * sched.length + 0.03, w: SCW, h: 0.18, fontSize: 7.5, color: C.accent5 });

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
footer("출처: CMS MKT | 2026 4Q · [4Q] Items통합 B~H열 에듀토크(10/6 기준) / 12월 모객 로드맵(9/9, Claude Code 작업) · D-7 리마인드일(10/28)은 방송일 기준 산출");

s.addNotes("[4Q] Items통합에서 담당자가 최세종인 아이템(에듀토크)만 정리. 협업 업무(주제 기획·패널: 홍혜숙/석다혜, 광고 소재·특전 제작: 최유나, 비주얼: 공혜영)는 일정 의존 관계로만 언급. 예산 금액은 시트에 없어 회의 전 기입 필요.");

// ================= 4. 에듀토크 상세 ① 준비 일정 =================
pres.addSection({ title: "에듀토크 상세" });
const eduTitle = (sub) => {
  s = pres.addSlide({ masterName: "ONE_PAGER", sectionTitle: "에듀토크 상세" });
  s.addText([{ text: "에듀토크 상세" }, { text: "   " + sub, options: { fontSize: 13, bold: false } }], { placeholder: "title" });
  T("담당 최세종 · 방송 11/4(수)", { x: 9.7, y: 0.28, w: 3.23, h: 0.5, align: "right", valign: "middle", color: C.accent5 });
};
eduTitle("① 준비 일정 — 24개 업무와 크리티컬 패스");

// [워크스트림, 날짜(M/D 또는 null), 업무, 담당, 상태(done|doing|todo), 크리티컬]
const EDU = [
  ["기획·콘텐츠", "9/22", "송출 채널 선정 (이병훈TV)", "-", "done", 0],
  ["기획·콘텐츠", "10/6", "주제 기획", "홍혜숙/석다혜", "doing", 1],
  ["기획·콘텐츠", "10/6", "패널 확정", "홍혜숙/석다혜", "doing", 1],
  ["기획·콘텐츠", "10/8", "특전 설정", "최세종", "todo", 1],
  ["기획·콘텐츠", "10/16", "특전 개발 제작", "최유나", "doing", 1],
  ["기획·콘텐츠", null, "특전 원고 작성 → 검수", "석다혜·이미현 → 홍혜숙", "todo", 0],
  ["기획·콘텐츠", "10/30", "원고 · 발표자료 개발", "홍혜숙/석다혜", "todo", 1],
  ["홍보·DM", "10/5", "홍보 채널 확정 (제휴실행안)", "최세종", "doing", 0],
  ["홍보·DM", "10/8", "채널별 광고 소재 기획", "최유나", "todo", 1],
  ["홍보·DM", null, "에듀토크 비주얼 기획", "공혜영", "todo", 0],
  ["홍보·DM", "10/16", "광고 소재 디자인 · DM 문안", "최유나 · 최세종", "todo", 1],
  ["홍보·DM", "10/19", "채널별 콘텐츠 게시", "최유나", "todo", 0],
  ["홍보·DM", "10/21", "DM 1차(D-14) · 2차(D-7)", "최세종", "todo", 1],
  ["랜딩·설문", "10/13", "랜딩 웹앱 구성 기획", "최세종/최유나", "todo", 0],
  ["랜딩·설문", "10/16", "설문 설계", "최세종", "todo", 0],
  ["랜딩·설문", "10/19", "랜딩 웹앱 개발", "최세종", "todo", 0],
  ["랜딩·설문", "10/21", "랜딩 릴리즈", "최세종", "todo", 1],
  ["방송 운영", "10/14", "라이브 채팅 대응방안 기획", "최세종", "todo", 0],
  ["방송 운영", "10/30", "대본 리딩·리허설 (11/2 최종 점검)", "홍혜숙", "todo", 1],
  ["방송 운영", "11/4", "방송 D-Day · 라이브 채팅 대응", "최세종", "todo", 1],
];
{
  const GX = 0.4, GW = 8.35, LBL = 3.55, AX = GX + LBL + 0.1, AW = GW - LBL - 0.35;
  const E0 = new Date(2026, 8, 20), E1 = new Date(2026, 10, 8);
  const ex = (md) => { const [m, d] = md.split("/").map(Number); return AX + ((new Date(2026, m - 1, d) - E0) / (E1 - E0)) * AW; };
  const top = 1.3, RH = 0.2, GH = 0.26;
  const streams = ["기획·콘텐츠", "홍보·DM", "랜딩·설문", "방송 운영"];
  const scol = { "기획·콘텐츠": C.accent2, "홍보·DM": C.accent1, "랜딩·설문": C.accent4, "방송 운영": C.accent3 };
  let y = top;
  const bodyH = streams.length * GH + EDU.length * RH;
  [["10/1", "10/1"], ["10/15", "10/15"], ["11/1", "11/1"]].forEach(([md, l]) => {
    T(l, { x: ex(md) - 0.3, y: top - 0.25, w: 0.6, h: 0.2, fontSize: 8, color: C.accent5, align: "center" });
    s.addShape(pres.shapes.LINE, { x: ex(md), y: top, w: 0, h: bodyH, line: { color: "D4D8DE", width: 0.5, dashType: "dash" } });
  });
  s.addShape(pres.shapes.LINE, { x: ex("10/7"), y: top, w: 0, h: bodyH, line: { color: HEX.accent6, width: 1.25 } });
  T("오늘", { x: ex("10/7") - 0.3, y: top - 0.25, w: 0.6, h: 0.2, fontSize: 8, bold: true, color: C.accent6, align: "center" });
  T("미정", { x: AX + AW + 0.02, y: top - 0.25, w: 0.35, h: 0.2, fontSize: 8, color: C.accent5, align: "center" });
  const crit = [];
  streams.forEach((st) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: GX, y: y + 0.03, w: GW, h: GH - 0.05, rectRadius: 0.04,
      fill: { color: scol[st], transparency: 85 }, line: { type: "none" } });
    T(st, { x: GX + 0.1, y, w: 2, h: GH, fontSize: 9, bold: true, color: scol[st], valign: "middle" });
    y += GH;
    EDU.filter((r) => r[0] === st).forEach(([, d, task, who, stt, cr]) => {
      const mine = who.includes("최세종");
      T([{ text: (d || "—") + "  ", options: { bold: true, color: C.text2 } }, { text: task, options: { bold: mine } },
        { text: "  " + who, options: { fontSize: 7.5, color: mine ? HEX.accent3 : C.accent5, bold: mine } }],
        { x: GX + 0.1, y, w: LBL, h: RH, fontSize: 8, valign: "middle" });
      s.addShape(pres.shapes.LINE, { x: AX, y: y + RH / 2, w: AW + 0.3, h: 0, line: { color: "EEF0F3", width: 0.5 } });
      const cx = d ? ex(d) : AX + AW + 0.2;
      if (cr && d) crit.push([cx, y + RH / 2]);
      s.addShape(pres.shapes.OVAL, { x: cx - 0.06, y: y + RH / 2 - 0.06, w: 0.12, h: 0.12,
        fill: { color: stt === "done" ? HEX.accent1 : stt === "doing" ? HEX.accent2 : "FFFFFF" },
        line: { color: cr ? HEX.accent6 : HEX.dk2, width: cr ? 1.5 : 0.75 } });
      y += RH;
    });
  });
  T([
    { text: "● 완료  ", options: { color: C.accent1, bold: true } }, { text: "● 진행중  ", options: { color: C.accent2, bold: true } },
    { text: "○ 예정  ", options: { color: C.text2, bold: true } }, { text: "◯ 빨간 테두리 = 크리티컬 패스  ", options: { color: C.accent6, bold: true } },
    { text: "굵은 글씨 = 최세종 담당", options: { color: C.accent3, bold: true } },
  ], { x: GX, y: top + bodyH + 0.08, w: GW, h: 0.22, fontSize: 8 });
}

// 크리티컬 패스·리스크 패널
{
  const PX = 8.95, PW = 3.98, PY = 0.98, PH = 5.4;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: PX, y: PY, w: PW, h: PH, rectRadius: 0.06, fill: { color: C.text2 }, line: { type: "none" } });
  T("크리티컬 패스", { x: PX + 0.2, y: PY + 0.08, w: PW - 0.4, h: 0.32, fontSize: 11, bold: true, color: C.background1, valign: "middle" });
  const path = [["10/6", "주제·패널 확정"], ["10/8", "특전 설정 · 소재 기획"], ["10/16", "특전 제작 · 소재 디자인 · DM 문안"], ["10/21", "랜딩 릴리즈 · DM 1차"], ["10/30", "원고·발표자료 · 리허설"], ["11/4", "방송 D-Day"]];
  path.forEach(([d, t], i) => {
    const y = PY + 0.5 + i * 0.36;
    s.addShape(pres.shapes.OVAL, { x: PX + 0.25, y: y + 0.06, w: 0.16, h: 0.16, fill: { color: i === path.length - 1 ? HEX.accent6 : "FFFFFF" }, line: { color: HEX.accent6, width: 1.25 } });
    if (i < path.length - 1) s.addShape(pres.shapes.LINE, { x: PX + 0.33, y: y + 0.22, w: 0, h: 0.2, line: { color: HEX.accent6, width: 1 } });
    T([{ text: d + "  ", options: { bold: true, color: "FFFFFF" } }, { text: t, options: { color: "DCE3EE" } }], { x: PX + 0.55, y, w: PW - 0.75, h: 0.28, fontSize: 9, valign: "middle" });
  });
  T("리스크와 대응", { x: PX + 0.2, y: PY + 2.75, w: PW - 0.4, h: 0.3, fontSize: 10, bold: true, color: C.background1 });
  const risks = [
    "주제·패널이 10/6보다 늦어지면 특전 → 소재 → DM이 연쇄 지연 → 오늘 회의에서 주제 확정",
    "DM 1차(10/21)는 랜딩 릴리즈와 같은 날 → 10/19 개발 완료 후 신청 테스트 1회 필수",
    "특전 원고·비주얼 기획은 날짜 미정 → 10/13까지로 지정 제안",
    "10/12 주에 팀 전체 업무 39건 집중 → 디자인 요청 순서 사전 합의",
  ];
  T(risks.map((t, i) => ({ text: t, options: { bullet: { indent: 10 }, breakLine: i < risks.length - 1, paraSpaceAfter: 5 } })),
    { x: PX + 0.2, y: PY + 3.08, w: PW - 0.4, h: 2.25, fontSize: 8.5, color: "DCE3EE", valign: "top" });
}
footer("출처: [4Q] Items통합 B~H열 에듀토크 24개 업무(10/6 기준) · ‘원고·발표자료’, ‘소재 디자인·DM 문안’, ‘DM 1·2차’ 등 같은 날짜 업무는 한 줄로 묶음 · 크리티컬 패스는 시트 비고의 선후관계 기준");

// ================= 5. 에듀토크 상세 ② 방송 운영안 =================
eduTitle("② 방송 운영안 (60분 기준, 안)");
{
  // 개요 표
  head(0.4, 0.98, 4.1, C.text2, "방송 개요");
  const ov = [
    ["일시", "11/4(수) — 방송 시간 확정 필요"],
    ["채널", "이병훈TV 라이브 (컨셉 대표님 보고 후 조율)"],
    ["패널", "이병훈 소장 + CMS 발표자 (10/6 확정 예정)"],
    ["주제", "10/6 확정 예정 (홍혜숙·석다혜) — 아래 후보 참고"],
    ["특전", "주제 확정 후 설정(10/8) → 제작(10/16)"],
    ["리허설", "10/30 대본 리딩·리허설 → 11/2 최종 점검"],
  ];
  s.addTable(ov.map(([k, v]) => [{ text: k, options: { bold: true, color: C.text2, fill: { color: "EEF1F5" } } }, { text: v }]),
    { x: 0.4, y: 1.38, w: 4.1, colW: [0.75, 3.35], rowH: 0.3, fontSize: 8.5, fontFace: THEME.bodyFontFace, color: C.text1,
      valign: "middle", margin: [0, 0.06, 0, 0.06], border: { type: "solid", pt: 0.5, color: "D4D8DE" } });

  // 주제 후보
  head(0.4, 3.35, 4.1, C.accent2, "주제 후보 (안)", "최종 선정은 주제 기획 담당");
  const topics = [
    ["A", "2027 특목·자사고 입시 변화와 겨울방학 수학 로드맵", "특사모·자사고 설명회와 연결"],
    ["B", "초·중 수학, 겨울방학 3개월이 1년을 결정한다", "FIT·MATHMILE 체험수업과 연결"],
    ["C", "이병훈 소장이 답하는 우리 아이 수학 진로 Q&A", "채팅 참여·설문 응답 극대화"],
  ];
  topics.forEach(([k, t, d], i) => {
    const y = 3.8 + i * 0.75;
    s.addShape(pres.shapes.OVAL, { x: 0.45, y: y + 0.08, w: 0.32, h: 0.32, fill: { color: C.accent2 }, line: { type: "none" } });
    T(k, { x: 0.45, y: y + 0.08, w: 0.32, h: 0.32, fontSize: 10, bold: true, color: C.background1, align: "center", valign: "middle" });
    T([{ text: t, options: { bold: true, breakLine: true } }, { text: d, options: { fontSize: 8, color: C.accent5 } }],
      { x: 0.88, y, w: 3.6, h: 0.5, fontSize: 9, valign: "middle" });
  });

  // 큐시트 바
  const QX = 4.75, QW = 8.18;
  head(QX, 0.98, QW, C.text2, "큐시트 (안)", "분 단위 · ▼ 표시는 신청·설문 링크 노출(CTA) 시점");
  const cue = [
    ["오프닝", 5, C.accent5, "패널 소개 · 방송 안내"],
    ["세션 1 · 입시 트렌드", 15, C.accent4, "이병훈 소장"],
    ["세션 2 · 겨울방학 로드맵", 15, C.accent1, "CMS 발표자"],
    ["특전 안내", 5, C.accent6, "특전 수령 = 설문 응답"],
    ["실시간 Q&A", 15, C.accent2, "채팅 질문 큐레이션"],
    ["클로징", 5, C.accent3, "인뎁스·체험 신청 안내"],
  ];
  const BY = 2.02, BH = 0.62, unit = QW / 60;
  let cx = QX, t0 = 0;
  cue.forEach(([n, m, col, d]) => {
    s.addShape(pres.shapes.RECTANGLE, { x: cx, y: BY, w: m * unit, h: BH, fill: { color: col }, line: { color: "FFFFFF", width: 1.5 }, objectName: "cue-" + n });
    T(m >= 15 ? n : n.replace(" · ", "\n"), { x: cx + 0.05, y: BY, w: m * unit - 0.1, h: BH, fontSize: m >= 15 ? 9 : 7.5, bold: true, color: C.background1, align: "center", valign: "middle" });
    T(d, { x: cx, y: BY + BH + 0.05, w: m * unit, h: 0.36, fontSize: 7.5, color: C.accent5, align: "center", valign: "top" });
    T(String(t0), { x: cx - 0.2, y: BY - 0.24, w: 0.4, h: 0.18, fontSize: 7.5, color: C.accent5, align: "center" });
    cx += m * unit; t0 += m;
  });
  T("60", { x: cx - 0.2, y: BY - 0.24, w: 0.4, h: 0.18, fontSize: 7.5, color: C.accent5, align: "center" });
  [[35, "CTA ① 특전·설문"], [45, "CTA ② 상담 신청"], [57, "CTA ③ 인뎁스·체험"]].forEach(([m, l]) => {
    const x = QX + m * unit;
    T("▼", { x: x - 0.15, y: BY - 0.47, w: 0.3, h: 0.2, fontSize: 10, color: C.accent6, align: "center" });
    const lx = Math.min(x - 0.9, QX + QW - 1.8);
    T(l, { x: lx, y: BY - 0.66, w: 1.8, h: 0.2, fontSize: 7.5, bold: true, color: C.accent6, align: lx < x - 0.9 ? "right" : "center" });
  });

  // 라이브 채팅 운영
  head(QX, 3.25, 4.0, C.accent3, "라이브 채팅 운영", "최세종 · 10/14 대응방안 기획");
  const roles = [
    ["진행 보조", "공지 고정, 신청·설문 링크 3회 게시, 방송 흐름 안내"],
    ["질문 큐레이터", "질문을 유형별로 분류 → Q&A용 상위 질문 전달"],
    ["상담 연결", "개별 상담성 질문 → 상담 신청 링크로 유도, 연락처는 채팅에서 받지 않음"],
  ];
  roles.forEach(([r, d], i) => {
    const y = 3.7 + i * 0.62;
    T([{ text: r, options: { bold: true, color: C.accent3, breakLine: true } }, { text: d }], { x: QX + 0.1, y, w: 3.8, h: 0.52, fontSize: 8.5, valign: "top" });
  });
  T("사전 준비: 예상 질문 FAQ 20개 · 금지 표현(과장·비교 광고) 가이드 · 부정 이슈 발생 시 보고 라인", { x: QX + 0.1, y: 5.6, w: 3.8, h: 0.5, fontSize: 8, color: C.accent5, valign: "top" });

  // D-Day 체크리스트
  const KX = QX + 4.18, KW = QW - 4.18;
  head(KX, 3.25, KW, C.accent1, "D-Day 체크리스트");
  const ck = ["D-1 송출 테스트·링크 점검", "랜딩 신청 마감 시간·안내 문구", "특전 다운로드 링크·자동 메시지", "설문 응답 → 시트 저장 확인", "채팅 FAQ·링크 문구 사전 작성", "방송 종료 후 다시보기 공개 여부", "D+1 후속 알림톡 예약 발송"];
  T(ck.map((t, i) => ({ text: "☐  " + t, options: { breakLine: i < ck.length - 1, paraSpaceAfter: 4 } })), { x: KX + 0.1, y: 3.7, w: KW - 0.2, h: 2.5, fontSize: 9, valign: "top" });
}
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.4, y: 6.44, w: 12.53, h: 0.6,
  fill: { color: C.accent6, transparency: 88 }, line: { color: C.accent6, width: 0.75 }, rectRadius: 0.06 });
T("회의 결정 필요", { x: 0.55, y: 6.44, w: 1.3, h: 0.6, fontSize: 10.5, bold: true, color: C.accent6, valign: "middle" });
T("① 주제 A·B·C 중 선택(또는 수정)   ② 방송 시간대   ③ 큐시트 시간 배분·CMS 발표자   ④ 채팅 운영 인원(역할 3개)   ⑤ 다시보기 공개 여부·기간",
  { x: 1.85, y: 6.44, w: 10.95, h: 0.6, fontSize: 9, valign: "middle" });
footer("시트 기준: 송출 채널·패널·주제 일정, 특전·리허설 일정, 라이브 채팅 대응 업무 · 큐시트·주제 후보·채팅 역할·체크리스트는 회의 논의용 제안(안)");

// ================= 6. 에듀토크 상세 ③ 모객 시퀀스·전환 설계 =================
eduTitle("③ 모객 시퀀스 · 신청 이후 전환 설계");
{
  head(0.4, 0.98, 12.53, C.text2, "메시지·홍보 시퀀스", "● 시트 확정   ○ 추가 제안");
  const S0 = new Date(2026, 9, 17), S1 = new Date(2026, 10, 22);
  const AX = 0.7, AW = 11.9;
  const sx = (md) => { const [m, d] = md.split("/").map(Number); return AX + ((new Date(2026, m - 1, d) - S0) / (S1 - S0)) * AW; };
  const LY = 2.5;
  s.addShape(pres.shapes.LINE, { x: AX, y: LY, w: AW, h: 0, line: { color: HEX.dk2, width: 1.5 } });
  const live = sx("11/4");
  s.addShape(pres.shapes.RECTANGLE, { x: live - 0.04, y: LY - 0.22, w: 0.08, h: 0.44, fill: { color: HEX.accent6 }, line: { type: "none" } });
  // [날짜, 제목, 설명, 시트확정, 위(1)/아래(0), 단(1|2)]
  const seq = [
    ["10/19", "콘텐츠 게시·광고", "채널별 순차 게시", 1, 1, 1],
    ["10/21", "랜딩 오픈·DM 1차", "D-14 예열", 1, 0, 1],
    ["10/28", "DM 2차", "D-7 리마인드", 1, 1, 1],
    ["11/1", "D-3 알림톡", "특전 미리보기", 0, 0, 1],
    ["11/3", "D-1 알림톡", "시청 링크·시간", 0, 1, 2],
    ["11/4", "방송 당일", "2시간 전 링크", 0, 0, 2],
    ["11/5", "D+1 다시보기", "설문 미응답자 재안내", 0, 1, 1],
    ["11/6", "인뎁스 알림톡", "에듀토크 종료 후", 1, 0, 1],
    ["11/20", "체험수업 신청 마감", "MATHMILE·FIT", 1, 1, 1],
  ];
  seq.forEach(([d, t, sub, fixed, up, tier]) => {
    const stem = tier === 2 ? 0.8 : 0.36;
    const x = sx(d);
    s.addShape(pres.shapes.OVAL, { x: x - 0.08, y: LY - 0.08, w: 0.16, h: 0.16, fill: { color: fixed ? HEX.dk2 : "FFFFFF" }, line: { color: HEX.dk2, width: 1.25 } });
    s.addShape(pres.shapes.LINE, { x, y: up ? LY - 0.1 - stem : LY + 0.1, w: 0, h: stem, line: { color: "9AA1AD", width: 0.5 } });
    T([{ text: d + " " + t, options: { bold: true, breakLine: true } }, { text: sub, options: { color: C.accent5 } }],
      { x: x + 1.45 > 12.93 ? x - 1.4 : x - 0.05, y: up ? LY - 0.52 - stem : LY + 0.12 + stem, w: 1.45, h: 0.4, fontSize: 8,
        align: x + 1.45 > 12.93 ? "right" : "left", valign: up ? "bottom" : "top" });
  });
  T("12월 2주 입학테스트 →", { x: 11.2, y: LY + 0.48, w: 1.73, h: 0.4, fontSize: 8, bold: true, color: C.accent3, align: "right" });

  // 설문 → 세그먼트 라우팅
  const RY = 3.95;
  head(0.4, RY, 7.0, C.accent1, "설문 → 세그먼트 → 후속 경로", "설문 설계 10/16 · 최세종");
  T([{ text: "설문 문항 (안)", options: { bold: true, breakLine: true, color: C.accent1 } },
    { text: "1. 자녀 학년", options: { breakLine: true } }, { text: "2. 목표 학교 유형", options: { breakLine: true } },
    { text: "3. 현재 수학 학습 방식", options: { breakLine: true } }, { text: "4. 관심 프로그램", options: { breakLine: true } },
    { text: "5. 상담 희망 여부", options: { breakLine: true } }, { text: "6. 궁금한 점 (주관식)" }],
    { x: 0.5, y: RY + 0.45, w: 1.75, h: 2.1, fontSize: 8.5, valign: "top" });
  T("→", { x: 2.2, y: RY + 1.1, w: 0.3, h: 0.4, fontSize: 16, color: C.accent5, align: "center" });
  const segs = [
    ["상담 희망", "방송 후 24~48h 내 상담콜 → 입테 예약", C.accent6],
    ["특목·자사고 목표", "인뎁스 세미나 · 특사모 면접 컨설팅 안내", C.accent3],
    ["MATHMILE 관심", "MATHMILE 체험수업 (신청 ~11/20)", C.accent2],
    ["FIT 관심", "FIT 체험수업 (매쏠로지2 체험계정)", C.accent1],
    ["기타·미응답", "D+1 다시보기 + 아샘·교과 알림톡", C.accent5],
  ];
  segs.forEach(([k, v, col], i) => {
    const y = RY + 0.45 + i * 0.42;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 2.55, y, w: 1.55, h: 0.34, rectRadius: 0.05, fill: { color: col }, line: { type: "none" } });
    T(k, { x: 2.6, y, w: 1.45, h: 0.34, fontSize: 8.5, bold: true, color: C.background1, align: "center", valign: "middle" });
    T("▶  " + v, { x: 4.2, y, w: 3.15, h: 0.34, fontSize: 8.5, valign: "middle" });
  });

  // KPI 퍼널 (목표 기입)
  const FX = 7.6, FW = 5.33;
  head(FX, RY, FW, C.accent3, "성과 지표", "목표는 회의에서 확정");
  const kpi = [["노출·도달", "광고·채널 리포트"], ["랜딩 방문", "웹앱 방문 로그"], ["사전 신청", "신청 DB"], ["라이브 시청", "이병훈TV 통계"], ["설문 응답", "설문 시트"], ["후속 신청", "인뎁스·체험 신청"], ["입테 응시 · 등록", "입테 명단·등록"]];
  kpi.forEach(([k, src], i) => {
    const w = FW - 1.35 - i * 0.22, x = FX + (FW - 1.35 - w) / 2, y = RY + 0.45 + i * 0.3;
    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.26, fill: { color: C.accent3, transparency: 15 + i * 10 }, line: { type: "none" } });
    T([{ text: k, options: { bold: true } }, { text: "  " + src, options: { fontSize: 7.5 } }], { x, y, w, h: 0.26, fontSize: 8.5, color: C.background1, align: "center", valign: "middle" });
    T("목표 ______", { x: FX + FW - 1.3, y, w: 1.3, h: 0.26, fontSize: 8.5, color: C.text2, valign: "middle" });
  });
}
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.4, y: 6.55, w: 12.53, h: 0.48, fill: { color: C.background2 }, line: { type: "none" }, rectRadius: 0.06 });
T([{ text: "핵심  ", options: { bold: true, color: C.text2 } },
  { text: "에듀토크는 신청 DB를 모으는 것으로 끝나지 않음 → 설문 응답으로 관심사를 나눠 인뎁스·체험수업·특사모로 바로 넘기고, 상담 희망자는 48시간 안에 연락" }],
  { x: 0.55, y: 6.55, w: 12.25, h: 0.48, fontSize: 9.5, valign: "middle" });
footer("● 시트 확정: 콘텐츠 게시 10/19 · 랜딩·DM 1차 10/21 · DM 2차(D-7) · 인뎁스 알림톡(에듀토크 종료 후) · 체험수업 신청 관리 ~11/20 / ○ 제안: D-3·D-1·당일·D+1 알림톡, 설문 문항, 세그먼트 경로");

// pptxgenjs는 테마 글꼴을 라틴(영문)에만 넣고 동아시아(ea) 글꼴은 비워 둔다.
// 비어 있으면 한글이 PC 기본 글꼴로 대체되므로 테마의 ea 글꼴도 같은 글꼴로 채운다.
async function setEastAsianFont(file, head, body) {
  const fs = require("fs");
  const JSZip = require(require.resolve("jszip", { paths: [require.resolve("pptxgenjs")] }));
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const path = "ppt/theme/theme1.xml";
  let xml = await zip.file(path).async("string");
  const fill = (tag, face) => { xml = xml.replace(new RegExp("(<a:" + tag + ">[\\s\\S]*?)<a:ea typeface=\"[^\"]*\"", ""), "$1<a:ea typeface=\"" + face + "\""); };
  fill("majorFont", head);
  fill("minorFont", body);
  zip.file(path, xml);
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  await setEastAsianFont(OUT, THEME.headFontFace, THEME.bodyFontFace);
  console.log("wrote", OUT);
})();
