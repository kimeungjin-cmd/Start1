// 시험항목 자동선정 에이전트 — 부사장 성과보고
const path = require("path");
const D = require(path.join(__dirname, "../../.claude/skills/company-report-pptx/scripts/deck.js"));
const pres = D.createDeck(), TH = D.THEME, F = f => path.join(__dirname, f);
const SRC = "※ 검증: 냉장고·세탁기·에어컨·조리기기 각 10개, 총 40개 모델 / 과거 개발단계 선정 목록 대비";

// 1 한 장 요약
(() => {
  const { slide, boxes } = D.slide(pres, { title: "시험항목 자동선정 에이전트 성과",
    message: "40개 모델에서 사람 선정 시험을 모두 포함, 놓치던 연관 시험 10% 추가",
    sub: "과거 개발단계 선정 대비 110% 선정 / 초보자도 전문가 수준의 검증 항목 선정 가능",
    layout: "2-64", source: SRC,
    boxes: [{ title: "검증 결과", msg: "4개 제품군 40개 모델 전체에서 확인" },
            { title: "의미와 다음 단계", msg: "사람 동등을 넘어 '생각 못 한 시험'을 찾음" }],
    notes: "가정: 대상 40개 모델(4제품군×10), 과거 선정 대비 110%(기존 100% 포함 + 연관 10% 추가). 시간 단축 실측 값 미입력" });
  D.figureFit(pres, slide, boxes[0].inner, F("v1.png"), "bar");
  D.rows(pres, slide, boxes[1].inner, [
    ["동등", "과거에 사람이 선정한 시험을\n빠짐없이 포함"],
    ["그 이상", "전원·열·노이즈 등 연관 영향\n시험을 10% 더 선정", true],
    ["시간", "선정·사례 검토 ○시간 → ○분\n[확인 필요]"],
    ["요청", "추가 시험 타당성 검토 후\n전 제품군 확대 적용"],
  ], { lw: 0.95 });
})();

// 2 배경과 목적
(() => {
  const { slide, boxes } = D.slide(pres, { title: "과제 배경과 목적",
    message: "시험항목 선정이 담당자 경험에 좌우되는 문제를 에이전트로 해결",
    sub: "변경점 하나를 넣으면 검증할 시험항목과 과거 유사 문제 사례를 함께 제시",
    layout: "2-55", source: "※ 과제 목적·범위: 과제 정의서 기준",
    boxes: [{ title: "기존 문제", msg: "같은 변경점도 사람마다 선정이 달라짐" },
            { title: "에이전트가 하는 일", msg: "최종 선정·판정은 사람, 누락 방지는 에이전트" }],
    notes: "과제 목적: 초보자도 전문가 수준 리스크 분석으로 검증 항목 선정 / 변경점 관련 과거 유사 문제 사례를 빠르게 가져와 검증 시 유의" });
  D.rows(pres, slide, boxes[0].inner, [
    ["경험 의존", "숙련자는 연관 영향까지 보지만\n경험이 적으면 직접 영향만 봄"],
    ["사례 검색", "과거 유사 불량을 찾는 데\n시간이 들고 사람마다 범위가 다름"],
    ["누락 리스크", "빠진 시험은 후속 단계·시장에서\n불량으로 드러남", true],
  ], { lw: 1.2 });
  const [a, b] = D.splitV(boxes[1].inner, [0.65, 1.1], 0.2);
  D.flow(pres, slide, a, [
    { title: "변경점 입력", desc: "부품·회로·S/W", state: "done" },
    { title: "리스크 분석", desc: "직접 + 연관 영향", state: "now" },
    { title: "담당자 확정", desc: "선정·판정", state: "plan" },
  ]);
  D.cards(pres, slide, b, [
    { tag: "결과①", title: "검증할 시험항목", body: "리스크 근거와 함께 목록으로 제시" },
    { tag: "결과②", title: "과거 유사 문제 사례", body: "같은 변경에서 났던 불량과 유의점" },
  ]);
})();

// 3 검증 방법과 결과
(() => {
  const { slide, boxes } = D.slide(pres, { title: "검증 방법과 결과",
    message: "과거 개발단계 선정 목록과 같은 변경점으로 1:1 비교",
    sub: "새 정답을 만들지 않고 실제 개발에 쓴 선정 결과를 기준으로 삼음",
    layout: "2-55", source: SRC,
    boxes: [{ title: "검증 방법", msg: "같은 입력, 같은 기준으로 비교" },
            { title: "결과", msg: "기존 시험 100% 포함 + 연관 10% 추가" }],
    notes: "측정 순서: 대상 선정 → 과거 목록 수집 → 에이전트 실행 → 항목 이름 대응표 → 산출" });
  D.steps(pres, slide, boxes[0].inner, [
    ["대상", "4개 제품군 × 10개 모델 = 40개 모델"],
    ["기준", "각 모델 개발단계에서 실제 선정한 시험 목록"],
    ["실행", "같은 변경점 내용을 에이전트에 입력"],
    ["대조", "이름이 다른 같은 시험은 대응표로 맞춤", true],
  ]);
  const [a, b] = D.splitV(boxes[1].inner, [1, 1], 0.15);
  D.kpis(pres, slide, a, [{ num: "40", unit: "개", label: "검증 모델" }, { num: "100", unit: "%", label: "기존 시험 포함" }]);
  D.kpis(pres, slide, b, [{ num: "+10", unit: "%", label: "추가 연관 시험", red: true }, { num: "○", unit: "분", label: "선정 시간 [확인 필요]" }]);
})();

// 4 연관 시험 [맞춤]
(() => {
  const { slide, boxes } = D.slide(pres, { title: "추가로 찾은 연관 시험",
    message: "사람은 직접 영향을 보고, 에이전트는 연결된 부위까지 따라간다",
    sub: "+10%는 변경 부품과 전원·기판·신호·S/W를 공유하는 부위의 시험",
    layout: "1", source: "※ 연결 고리 유형은 대표 예시 / 사례 표의 내용은 실제 검증 사례로 교체 필요",
    boxes: [{ title: "연관 시험이 나오는 구조와 사례", msg: "'생각 못 했던 시험'은 연결 고리를 따라간 결과" }],
    notes: "연관 시험 = 과거 선정에 없고 에이전트만 선정한 항목. 사례 3건은 [확인 필요]" });
  const [a, b] = D.splitV(boxes[0].inner, [2.2, 1.55], 0.15);
  D.figureFit(pres, slide, a, F("v2.png"), "diagram");
  const w = b.w;
  D.table(pres, slide, b, ["제품군", "변경점", "기존 선정", "에이전트 추가 연관 시험", "연결 고리"], [
    ["[확인 필요]", "[확인 필요]", "[확인 필요]", "[확인 필요]", "전원 공유 등"],
    ["[확인 필요]", "[확인 필요]", "[확인 필요]", "[확인 필요]", "과거 유사 불량 등"],
  ], { colW: [1.4, 2.6, 2.4, w - 8.0, 1.6], rowH: 0.5, fontSize: 11 });
})();

// 5 한계와 확대 계획
(() => {
  const { slide, boxes } = D.slide(pres, { title: "남은 확인과 확대 계획",
    message: "추가 시험의 타당성 검토 후 전 제품군 현업 적용으로 확대",
    sub: "현재 결과는 과거 선정 대비 비교 / 실제 불량 감소 효과는 적용 후 측정",
    layout: "2-64", source: "※ 일정·담당은 [확인 필요]",
    boxes: [{ title: "확대 일정", msg: "검토 → 시범 → 확대 → 효과 측정" },
            { title: "현재 결과의 범위", msg: "말할 수 있는 것과 아직인 것을 구분" }],
    notes: "놓친 불량 근거 데이터 현재 없음 → 적용 후 선정 이력과 불량 연결 기록" });
  D.timeline(pres, slide, boxes[0].inner, [
    { date: "1단계", title: "추가 시험 검토", desc: "+10% 숙련자 확인", red: true },
    { date: "2단계", title: "현업 시범", desc: "신규 개발 모델 적용" },
    { date: "3단계", title: "전 제품군 확대", desc: "사례 DB 보강 병행" },
    { date: "4단계", title: "효과 측정", desc: "선정 이력 ↔ 불량 연결", red: true },
  ]);
  D.rows(pres, slide, boxes[1].inner, [
    ["확인됨", "과거 선정 시험 100% 포함\n(40개 모델)"],
    ["검토 중", "추가 10%가 모두 필요한\n시험인지 → 숙련자 확인", true],
    ["측정 예정", "불량 감소 효과는\n적용 후 측정"],
    ["전제", "사례 DB에 없는 과거 불량은\n찾을 수 없음"],
  ], { lw: 1.05 });
})();

// 유첨 1 측정 정의
(() => {
  const { slide, boxes } = D.slide(pres, { title: "유첨1. 지표 정의와 측정 기준",
    message: "110%는 '기존 시험 100% 포함 + 연관 시험 10% 추가'의 합",
    sub: "커버율 자체는 100%가 최대 → 포함률과 추가율로 나눠 표기", layout: "1", source: SRC,
    boxes: [{ title: "지표 정의", msg: "본문 숫자는 모두 이 정의로 계산" }],
    notes: "본문 1·3장 근거" });
  const A = boxes[0].inner;
  const y = D.table(pres, slide, A, ["지표", "계산식", "결과", "비고"], [
    ["기존 시험 포함률", "에이전트도 선정한 기존 시험 수 / 과거 선정 시험 수", "100%", "사람 동등 기준"],
    ["연관 시험 추가율", "에이전트만 선정한 시험 수 / 과거 선정 시험 수", "+10%", "타당성 검토 중"],
    ["선정 규모", "에이전트 선정 시험 수 / 과거 선정 시험 수", "110%", "포함률 + 추가율"],
    ["선정 시간", "변경점 입력 → 목록 완성 (사람 검토 포함)", "[확인 필요]", "과거 기록 없으면 추정치"],
  ], { colW: [2.0, 5.6, 1.4, A.w - 9.0], rowH: 0.55, fontSize: 11.5, hi: 1 });
  D.T(slide, "이름이 달라도 같은 시험이면 같은 항목으로 셈 — 대응표는 별도 관리", { x: A.x, y: y + 0.15, w: A.w, h: 0.4, fontSize: 11.5, color: TH.gray });
})();

D.save(pres, path.join(__dirname, "시험항목자동선정_성과보고.pptx"), { reserveCover: true }).then(() => console.log("saved"));
