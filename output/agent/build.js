// 시험항목 자동선정 에이전트 — 과제 정의서 (3장)
const path = require("path");
const D = require(path.join(__dirname, "../../.claude/skills/company-report-pptx/scripts/deck.js"));
const pres = D.createDeck(), TH = D.THEME;
const SRC = "※ 과제 목적·성과 기준: 과제 담당자 정의 / [확인 필요] 값은 파트장 협의 후 확정";

// 1 과제 정의
(() => {
  const { slide, boxes } = D.slide(pres, { title: "시험항목 자동선정 에이전트 과제 정의",
    message: "초보자도 전문가 수준으로 리스크를 분석해 검증 항목을 선정하게 함",
    sub: "변경점 입력 → 리스크 분석·시험항목 선정 + 과거 유사 문제 사례 제시",
    layout: "2-64", source: SRC,
    boxes: [{ title: "에이전트가 하는 일", msg: "변경점 하나로 두 결과를 함께 냄" },
            { title: "과제 범위", msg: "선정과 사례 검토까지, 판정은 사람" }],
    notes: "사용자 제공: 목적(초보자도 전문가 수준 리스크 분석, 검증 항목 선정 / 변경점 관련 과거 유사 문제 사례를 빠르게 가져와 검증 시 유의), 성과기준(사람 동등 이상 시험항목 선정, 과거 유사사례 검토로 시간 단축·정확도 확보)" });
  const [a, b] = D.splitV(boxes[0].inner, [1, 1.05], 0.25);
  D.flow(pres, slide, a, [
    { title: "변경점 입력", desc: "부품·회로·S/W\n변경 내용", state: "done" },
    { title: "리스크 분석", desc: "변경으로 생길\n고장 모드 도출", state: "now" },
    { title: "결과 2종", desc: "시험항목 목록\n+ 과거 유사 사례", state: "now" },
    { title: "담당자 검토", desc: "최종 선정·\n판정은 사람", state: "plan" },
  ]);
  D.cards(pres, slide, b, [
    { tag: "결과①", title: "검증할 시험항목", body: "리스크 근거와 함께 선정 → 초보자도 누락 없이" },
    { tag: "결과②", title: "과거 유사 문제 사례", body: "같은 변경에서 났던 불량 → 검증 시 유의점" },
  ]);
  D.rows(pres, slide, boxes[1].inner, [
    ["사용자", "시험항목을 선정하는 담당자\n(특히 경험이 적은 담당자)"],
    ["입력", "변경점 내용 [확인 필요: 입력 양식]"],
    ["포함", "리스크 분석, 시험항목 선정,\n유사 사례 검색·요약"],
    ["제외", "시험 판정, 시험 실행 자동화", true],
  ], { lw: 0.95 });
})();

// 2 성과 지표 (A안)
(() => {
  const { slide, boxes } = D.slide(pres, { title: "성과 지표와 비교 기준",
    message: "과거 개발단계에서 사람이 선정한 항목과 비교해 커버율·시간으로 본다",
    sub: "A안: 실제 개발에 쓴 선정 결과를 기준으로 비교 → 새 정답 작업 없이 측정",
    layout: "2-64", source: SRC,
    boxes: [{ title: "지표 정의", msg: "커버율로 '동등', 시간으로 '단축'을 봄" },
            { title: "목표", msg: "목표 값은 파트장 협의로 확정" }],
    notes: "A안 확정(①과거 선정 커버율 ③시간). B안(차이 항목 숙련자 검토)은 확장 단계로 보류" });
  const A = boxes[0].inner;
  const y = D.table(pres, slide, A, ["지표", "정의", "용도"], [
    ["선정 커버율", "과거 사람이 선정한 항목 중\n에이전트도 선정한 비율", "사람 동등 여부"],
    ["선정 소요 시간", "변경점 입력 → 시험항목\n목록 완성까지", "시간 단축"],
    ["사례 검토 시간", "과거 유사 문제 사례를\n찾고 읽는 데 든 시간", "시간 단축"],
    ["추가 선정 (참고)", "에이전트만 선정한 항목 건수\n(필요 여부 판정은 안 함)", "차이 목록으로 첨부"],
  ], { colW: [1.75, A.w - 3.55, 1.8], rowH: 0.68, fontSize: 11.5, hi: 0 });
  D.T(slide, "과거 선정이 정답은 아님 → 서로 다르게 선정한 항목은 차이 목록으로 남김",
    { x: A.x, y: y + 0.12, w: A.w, h: A.y + A.h - y - 0.12, fontSize: 11.5, bold: true, color: TH.red, valign: "middle" });
  const [a, b] = D.splitV(boxes[1].inner, [1, 1], 0.15);
  D.kpis(pres, slide, a, [{ num: "○○", unit: "%", label: "커버율 목표\n[확인 필요]" }, { num: "○○", unit: "%", label: "시간 단축 목표\n[확인 필요]" }]);
  D.rows(pres, slide, b, [
    ["동등 판단", "커버율이 목표 이상"],
    ["단축 판단", "선정 + 사례 검토 시간 합계로 비교"],
    ["비교 대상", "같은 변경점의 과거 개발단계 선정"],
  ], { lw: 1.1 });
})();

// 3 측정 방법 (A안)
(() => {
  const { slide, boxes } = D.slide(pres, { title: "성과 측정 방법과 한계",
    message: "과거 변경점 N건에 에이전트를 돌려 당시 선정 목록과 맞춰 본다",
    sub: "대상 건수·일정은 [확인 필요] / 실제 불량 감소 효과는 지금부터 기록해 추후 측정",
    layout: "2-55", source: SRC,
    boxes: [{ title: "측정 순서", msg: "항목 이름 맞추기가 정확도를 좌우" },
            { title: "한계와 다음 단계", msg: "지금 말할 수 있는 범위를 먼저 밝힘" }],
    notes: "A안 측정 순서. 놓친 불량 지표는 현재 근거 데이터 없음 → 측정 계획으로 전환" });
  D.steps(pres, slide, boxes[0].inner, [
    ["대상 선정", "과거 변경점 N건 [확인 필요], 모델·변경 유형 섞어서"],
    ["과거 목록 수집", "당시 선정한 시험항목 목록과 소요 시간 기록"],
    ["에이전트 실행", "같은 변경점 내용을 입력해 목록·사례 출력"],
    ["항목 이름 맞추기", "같은 시험인데 이름이 다른 항목 → 대응표로 정리", true],
    ["결과 산출", "커버율, 시간, 차이 항목 목록"],
  ]);
  const [a, b] = D.splitV(boxes[1].inner, [0.8, 1.4], 0.2);
  D.cards(pres, slide, a, [
    { tag: "보고 예", title: "\"과거 선정 대비 커버율 ○○%\"", body: "선정 + 사례 검토 ○시간 → ○분 / 차이 항목 ○건 첨부" },
  ]);
  D.rows(pres, slide, b, [
    ["과거 시간", "기록이 없으면 담당자 추정치로 명시"],
    ["불량 감소", "근거 데이터 없음 → 지금부터 기록\n○월 첫 비교 [확인 필요]", true],
    ["확장 (B안)", "차이 항목을 숙련자가 검토해\n'사람보다 나은 부분'을 측정"],
  ], { lw: 1.1 });
})();

D.save(pres, path.join(__dirname, "시험항목자동선정_과제정의서.pptx"), { reserveCover: true }).then(() => console.log("saved"));
