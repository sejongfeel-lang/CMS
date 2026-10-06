// 겨울학기 모객 전략안 1페이지 PPT 생성 스크립트
// 실행: NODE_PATH=<pptxgenjs가 설치된 node_modules> node scripts/build_winter_onepager.js <출력.pptx> <apply_theme.js 경로>
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.argv[3]);
const OUT = process.argv[2] || "겨울학기_모객전략_1p.pptx";

const THEME = {
  name: "CMS Winter",
  headFontFace: "Malgun Gothic",
  bodyFontFace: "Malgun Gothic",
  colors: {
    dk1: "1B2130", lt1: "FFFFFF", dk2: "1F2A44", lt2: "EEF1F5",
    accent1: "2A7A72", // 인지·DB (teal)
    accent2: "A8640F", // 체험·관계 (amber)
    accent3: "8A3F5C", // 응시·등록 (plum)
    accent4: "3C5A91", accent5: "5B6472", accent6: "C8553D",
    hlink: "3C5A91", folHlink: "8A3F5C",
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "겨울학기 모객 전략안";
const C = pres.SchemeColor;

pres.defineSlideMaster({
  title: "ONE_PAGER",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.4, y: 0.28, w: 8.6, h: 0.5,
      fontSize: 22, bold: true, color: C.text2, margin: 0, align: "left", valign: "middle" }, text: "" } },
  ],
});

pres.addSection({ title: "겨울학기 모객 전략" });
const s = pres.addSlide({ masterName: "ONE_PAGER", sectionTitle: "겨울학기 모객 전략" });
const T = (text, opts) => s.addText(text, Object.assign({ isTextBox: true, margin: 0, fontSize: 9, color: C.text1 }, opts));

s.addText([{ text: "겨울학기 모객 전략안" }, { text: "   4Q 확정 아이템 실행 방향 + 추가 제안", options: { fontSize: 13, bold: false } }], { placeholder: "title" });
T("10/7(수) 11:00 대회의실 · 마케팅팀", {
  x: 9.1, y: 0.28, w: 3.83, h: 0.5, align: "right", valign: "middle", color: C.accent5 });
T([
  { text: "핵심 방향  ", options: { bold: true, color: C.text2 } },
  { text: "10월 인지·DB 확보 → 11월 체험·관계 형성 → 12월 입테 응시·등록 전환 — 모든 아이템의 DB를 하나의 퍼널로 연결" },
], { x: 0.4, y: 0.84, w: 12.53, h: 0.32, fontSize: 10.5, valign: "middle" });

// ---------- 타임라인 (10/1 ~ 12/31) ----------
const X0 = 0.6, XW = 12.1, DAYS = 91;
const xd = (d) => X0 + (d / DAYS) * XW; // d = 10/1 기준 일수
const bars = [[0, 31, C.accent1, "10월"], [31, 61, C.accent2, "11월"], [61, 92, C.accent3, "12월"]];
for (const [a, b, col, lbl] of bars) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: xd(a), y: 1.64, w: xd(b) - xd(a) - 0.04, h: 0.2,
    fill: { color: col }, line: { type: "none" }, rectRadius: 0.05, objectName: "timeline-" + lbl });
  T(lbl, { x: xd(a) + 0.06, y: 1.64, w: 0.36, h: 0.2, fontSize: 8.5, bold: true, color: C.background1, valign: "middle" });
}
const ms = [
  [5, "10/6 예산 확정·현장마케팅 취합", true, 2.4],
  [20, "10/21 메타광고·에듀토크 DM", false, 1.8],
  [25, "10/26 카페·체험수업 게시", true, 1.8],
  [34, "11/4 에듀토크 방송", false, 1.3],
  [39, "11/9 친구추천 오픈", true, 1.4],
  [44, "11/14 KMO 2차 현장", false, 1.6],
  [50, "11/20 체험수업 신청 마감", true, 2.3],
  [64, "12/4 친구추천 종료", false, 1.8],
  [69, "12/2주 입테 · 12/11 리워드 지급", true, 2.6],
  [78, "12/18 현장마케팅 정산", false, 1.7],
];
for (const [d, lbl, above, w] of ms) {
  const x = xd(d);
  s.addShape(pres.shapes.OVAL, { x: x - 0.06, y: 1.68, w: 0.12, h: 0.12,
    fill: { color: C.background1 }, line: { color: C.text2, width: 1.25 } });
  s.addShape(pres.shapes.LINE, { x, y: above ? 1.5 : 1.84, w: 0, h: 0.14, line: { color: C.accent5, width: 0.75 } });
  T(lbl, { x: x - 0.04, y: above ? 1.27 : 1.98, w, h: 0.22, fontSize: 8.5, color: C.text1, valign: "middle" });
}

// ---------- 확정 아이템: 퍼널 3단계 ----------
const COLS = [
  { x: 0.4, color: C.accent1, head: "① 인지·DB 확보", when: "10월~11월 초", items: [
    ["에듀토크 (이병훈TV 라이브)", "최세종", "학부모 입시 토크 라이브 11/4 · 특전·설문", "랜딩 DB → DM D-14/D-7 → 인뎁스·체험", "특전 제작·채널 홍보비"],
    ["메타(인스타) 광고", "최유나", "학부모 타깃 광고 · 앰플랜잇 대행 · 10/21~", "랜딩 신청폼 DB → 24h 효율 점검·개선", "광고비 + 대행 수수료"],
    ["인플루언서 마케팅", "최유나", "셀러 포스팅 + 특전자료 다운로드", "DB 폼·자동 메시지 → 입테 응시 리워드", "에이전시비·리워드·교재 발송"],
    ["당근마켓 · 네이버 카페", "석다혜/이미현", "맘카페·당근에 에듀토크·체험수업 노출", "체험수업 신청폼 DB (카페글 10/26)", "카페 집행비·당근 광고비"],
    ["KMO 2차 · 자사고 설명회", "김새미/홍혜숙", "판촉물 + 진학 가이드 QR 특전 배포", "QR 유입 DB → 알림톡 리타깃", "판촉물·배포 인력·배송"],
  ]},
  { x: 3.22, color: C.accent2, head: "② 체험·관계 형성", when: "10월 말~11월", items: [
    ["인뎁스 세미나", "석다혜", "센터별 학부모 세미나 · 발표자료 10/29", "센터 블로그 10/23 → 현장 입테 예약", "홍보 이미지·센터 운영비"],
    ["아샘", "이미현", "초·중 진행 방식 10/6 · 자료 개발 10/20", "블로그 홍보 → 참석자 상담·입테 예약", "자료 제작비"],
    ["FIT · MATHMILE 체험수업", "석다혜/이미현", "매쏠로지2 체험계정 · 초등 기프트 특전", "랜딩·카페 신청 → 11/20까지 관리 → 등록 상담", "기프트·활동지·체험계정"],
    ["[특사모] 자사고 면접 컨설팅", "홍혜숙", "자사고 지원 예비 고1 면접 컨설팅", "참여자 → 겨울학기 특목·자사 과정", "별도 마스터시트 기준"],
  ]},
  { x: 6.04, color: C.accent3, head: "③ 응시·등록 전환", when: "11월~12월", items: [
    ["교과 프로그램 홍보 알림톡", "이미현/석다혜", "FIT·PL / GS·MATHMILE 문안 11/3", "보유 DB → 입테 예약·등록 (발송 김새미)", "알림톡 발송비"],
    ["친구추천 이벤트", "이로은", "재원생 추천 11/9~12/4 (CMS 공통)", "지인 입테 예약 → 리워드 12/11", "경품·리워드"],
    ["수강후기 프로모션", "최유나", "재원생 후기 이벤트 10/20~11/15(안)", "후기 → 광고·카페·블로그 신뢰 소재", "경품"],
    ["현장 주도 마케팅 지원", "최유나", "센터 BM 제안 취합 → 10/14 결정 공지", "센터별 지역 맞춤 모객 → 정산 12/18", "센터별 승인액"],
  ]},
];
const TOP = 2.32, COLW = 2.7, BODY_H = 3.98;
for (const col of COLS) {
  s.addShape(pres.shapes.RECTANGLE, { x: col.x, y: TOP, w: COLW, h: BODY_H,
    fill: { color: col.color, transparency: 92 }, line: { type: "none" }, objectName: "col-" + col.head });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: col.x, y: TOP, w: COLW, h: 0.38,
    fill: { color: col.color }, line: { type: "none" }, rectRadius: 0.06 });
  T([
    { text: col.head, options: { bold: true, fontSize: 11 } },
    { text: "  " + col.when, options: { fontSize: 8.5 } },
  ], { x: col.x + 0.1, y: TOP, w: COLW - 0.2, h: 0.38, color: C.background1, valign: "middle" });

  const runs = [];
  col.items.forEach(([name, pm, how, link, budget], i) => {
    runs.push({ text: name, options: { bold: true, fontSize: 9.5, color: C.text2, paraSpaceBefore: i ? 7 : 0, breakLine: false } });
    runs.push({ text: "  " + pm, options: { fontSize: 8, color: C.accent5, breakLine: true } });
    runs.push({ text: how, options: { fontSize: 8.5, breakLine: true } });
    runs.push({ text: "→ " + link, options: { fontSize: 8.5, bold: true, color: col.color, breakLine: true } });
    runs.push({ text: "예산 · " + budget, options: { fontSize: 8, color: C.accent5, breakLine: i < col.items.length - 1 } });
  });
  T(runs, { x: col.x + 0.1, y: TOP + 0.46, w: COLW - 0.2, h: BODY_H - 0.52, valign: "top", lineSpacingMultiple: 0.98 });
}

// ---------- 추가 제안 (12월 모객 로드맵 기반) ----------
const PX = 8.92, PW = 4.01;
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: PX, y: TOP, w: PW, h: BODY_H,
  fill: { color: C.text2 }, line: { type: "none" }, rectRadius: 0.06, objectName: "proposal-panel" });
T([
  { text: "✚ 추가 모객 전략 제안", options: { bold: true, fontSize: 11 } },
  { text: "  12월 모객 로드맵 기반", options: { fontSize: 8.5, color: "CADCFC" } },
], { x: PX + 0.15, y: TOP, w: PW - 0.3, h: 0.38, color: C.background1, valign: "middle" });
const props = [
  ["A. 입학테스트 통합 퍼널·CRM", "자사고 모객 허브의 퍼널·CRM 구조를 이식 → 전 아이템 DB를 예약→응시→등록 단계로 통합 추적"],
  ["B. 노쇼·이탈 방지 자동 알림톡", "입테 D-14/7/3/1 리마인더 + 결과 통보 후 24~48h 골든타임 상담콜"],
  ["C. 아웃바운드 콜 스코어링", "미전환 DB 우선순위 자동 산출 → 알림톡 예열 → 표준 스크립트 콜 → CRM 기록"],
  ["D. AI 모의면접 체험 이벤트", "특목고 면접대비 웹앱을 리드마그넷으로 개방 (특사모·자사고 설명회 연계)"],
  ["E. 추천·신뢰 강화", "형제·자매 혜택을 친구추천에 결합 · 네이버 플레이스 리뷰 캠페인(수강후기 연계)"],
  ["F. 데이터 기반 타깃·콘텐츠", "지역경쟁분석 우선 세그먼트로 메타·당근 타깃 설정 · Montelena 여론 데이터로 시기별 콘텐츠 초안"],
];
const pruns = [];
props.forEach(([h, d], i) => {
  pruns.push({ text: h, options: { bold: true, fontSize: 9.5, color: "FFFFFF", paraSpaceBefore: i ? 6 : 0, breakLine: true } });
  pruns.push({ text: d, options: { fontSize: 8.5, color: "DCE3EE", breakLine: i < props.length - 1 } });
});
T(pruns, { x: PX + 0.15, y: TOP + 0.48, w: PW - 0.3, h: BODY_H - 0.56, valign: "top" });

// ---------- 회의 결정 필요 ----------
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.4, y: 6.44, w: 12.53, h: 0.6,
  fill: { color: C.accent6, transparency: 88 }, line: { color: C.accent6, width: 0.75 }, rectRadius: 0.06, objectName: "decisions" });
T("회의 결정 필요", { x: 0.55, y: 6.44, w: 1.3, h: 0.6, fontSize: 10.5, bold: true, color: C.accent6, valign: "middle" });
T("① 12월 입학테스트 일정·회차 및 센터별 목표(예약/응시/등록)   ② 아이템별 예산 금액 기입·합계   ③ 인뎁스·입학테스트 공통 예약폼 여부   ④ 당근마켓 품의 에듀토크 통합 여부   ⑤ 신규 알림톡 템플릿 카카오 심사(1~2영업일) 일정 반영   ⑥ 추가 제안(A~F) 채택·담당 지정",
  { x: 1.85, y: 6.44, w: 10.95, h: 0.6, fontSize: 9, valign: "middle" });
T("출처: CMS MKT | 2026 4Q · [4Q] Items통합 B~H열(10/6 기준) / 12월 모객 로드맵(9/9, Claude Code 작업) · 예산은 항목만 표기, 금액은 담당자 기입",
  { x: 0.4, y: 7.1, w: 12.53, h: 0.22, fontSize: 7.5, color: C.accent5 });

s.addNotes("4Q Items통합 B열 확정 아이템을 퍼널 3단계(인지·DB → 체험·관계 → 응시·등록)로 재배치하고, 각 아이템의 대상·방식, 모객 연결, 예산 항목을 정리. 예산 금액은 시트에 기재되어 있지 않아 담당자 기입 필요. 오른쪽은 12월 모객 로드맵에서 도출한 추가 제안.");

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
