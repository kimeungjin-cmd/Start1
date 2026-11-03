// 부품 사용 견본 (가상 데이터) — 부품 호출법이 헷갈릴 때 '코드 쓰는 법'만 참고한다.
// 장 구성·문구·데이터를 따라 하지 않는다(내용은 원본과 스토리보드에서 나온다).
// 실행: node examples/catalog.js [출력.pptx] [그림폴더]   그림폴더에 diag.png(도식), eq/E1.png(수식)가 있으면 8장에 사용
const path = require("path"), fs = require("fs");
const D = require(path.join(__dirname, "..", "scripts", "deck.js"));
const out = process.argv[2] || "부품_견본.pptx", figDir = process.argv[3] || ".";
const pres = D.createDeck(); const SRC = "※ 견본용 가상 데이터";
const S = (title, message, sub, layout, boxes) => D.slide(pres, { title, message, sub, layout, boxes, source: SRC });
const L1 = ["접점 용량 부족", "돌입전류 과다", "제조 불량", "미확인"], V1 = [7, 3, 1, 1];

let r = S("① 막대 · 도넛", "막대로 순위를, 도넛으로 비중을", "chartBar / chartDonut", "2-55", [{ title: "원인별 건수", msg: "접점 용량 부족이 최다" }, { title: "비중", msg: "용량 계열 83%" }]);
const bc = D.chartBar(pres, r.slide, r.boxes[0].inner, { labels: L1, values: V1, unit: "건", highlight: [0], related: [1] });
D.annotate(pres, r.slide, r.boxes[0].inner, { ...bc.barEnd(0), text: "유력 원인", dx: 0.25, dy: 0.45, w: 1.2 });  // 막대 끝에 판단 붙이기
D.chartDonut(pres, r.slide, r.boxes[1].inner, { labels: L1, values: V1, center: "58%", centerSub: "7 / 12건", unit: "건" });

r = S("② 꺾은선 · 주석 · 숫자 카드", "추이는 꺾은선 + 기준선, 핵심 지점엔 주석", "chartLine / annotate / kpis", "2-73", [{ title: "월별 시장불량률", msg: "9월 대책 후 0.3%" }, { title: "핵심 숫자", msg: "기준 0.5% 아래" }]);
const lc = D.chartLine(pres, r.slide, r.boxes[0].inner, { labels: ["6월", "7월", "8월", "9월", "10월"], series: [{ name: "불량률", values: [0.8, 0.9, 0.7, 0.4, 0.3] }], unit: "%", target: 0.5 });
D.annotate(pres, r.slide, r.boxes[0].inner, { ...lc.point(3, 0.4), text: "9월 대책 적용", dx: -2.3, dy: 0.5 });  // 점 위치는 point(i, 값)으로 계산
let [a, b] = D.splitV(r.boxes[1].inner, [1, 1]);
D.kpis(pres, r.slide, a, [{ num: "0.3", unit: "%", label: "10월 불량률" }]);
D.kpis(pres, r.slide, b, [{ num: "-63", unit: "%", label: "6월 대비", red: true }]);

r = S("③ 단계 화살표 · 타임라인", "진행은 단계 화살표, 일정은 타임라인", "flow / timeline", "1", [{ title: "원인 확정까지", msg: "재현시험이 다음 결정 지점" }]);
[a, b] = D.splitV(r.boxes[0].inner, [1, 1.3], 0.3);
D.flow(pres, r.slide, a, [{ title: "배제 원인", desc: "제조 불량 배제", state: "done" }, { title: "유력 가설", desc: "접점 용량 부족", state: "now" }, { title: "확인 계획", desc: "재현시험 11/20", state: "plan" }, { title: "리스크 조치", desc: "대체 검토 11/30", state: "risk" }]);
D.timeline(pres, r.slide, b, [{ date: "9/5", title: "첫 접수" }, { date: "9/30", title: "회수 12건" }, { date: "10/28", title: "분석 완료" }, { date: "11/20", title: "재현시험", red: true }, { date: "11/30", title: "대체 결정", red: true }]);

r = S("④ 위치도 · 판정표", "선택지는 2×2 위치도, 근거는 판정표", "matrix / table(judge, hi)", "2-55", [{ title: "대체 부품 후보", msg: "B가 비용·여유 모두 유리" }, { title: "정격 대비 여유율", msg: "현행만 NG" }]);
D.matrix(pres, r.slide, r.boxes[0].inner, { xLabel: "비용", yLabel: "정격 여유", goodZone: "tl", goodText: "싸고 여유 큼", items: [{ name: "후보 A", x: 0.7, y: 0.75 }, { name: "후보 B", x: 0.3, y: 0.8, red: true, note: "추천" }, { name: "현행", x: 0.25, y: 0.25 }] });
let R = r.boxes[1].inner;
D.table(pres, r.slide, R, ["부품", "정격", "여유율", "판정"], [["현행", "16A", "0.9배", "NG"], ["후보 A", "25A", "1.4배", "OK"], ["후보 B", "30A", "1.7배", "OK"]], { colW: [1.4, 1.3, 1.3, R.w - 4.0], hi: 2, judge: 3 });

r = S("⑤ 카드 · 체크리스트", "결론·추천은 카드, 해야 할 일은 체크리스트", "cards / checklist", "2-55", [{ title: "결론", msg: "추천 1 · 대안 1 · 보류 1" }, { title: "다음 확인", msg: "필수 3 + 선택 1" }]);
D.cards(pres, r.slide, r.boxes[0].inner, [
  { tag: "추천", red: true, title: "후보 B 30A 릴레이", big: "여유율 1.7배", body: "정격 여유·비용 모두 유리, 단 신뢰성 시험 필요" },
  { tag: "대안", title: "후보 A 25A 릴레이", big: "여유율 1.4배", body: "여유는 충분하나 비용 높음" },
  { tag: "보류", title: "현행 유지 + NTC 상향", body: "돌입전류만 낮추는 방안, 효과 검증 전" }]);
D.checklist(pres, r.slide, r.boxes[1].inner, [["재현시험 결과 확인", "11/20, 회로품질", "필수"], ["후보 B 신뢰성 시험", "개폐 수명·온도", "필수"], ["원가 영향 확인", "구매 협의", "필수"], ["횡전개 대상 모델 목록", "같은 릴레이 사용 모델", "선택"]]);

r = S("⑥ 행 목록 · 단계 목록", "속성 설명은 행 목록, 순서가 있는 일은 단계 목록", "rows / steps", "2-55", [{ title: "검토 조건", msg: "판정 기준과 범위" }, { title: "진행 순서", msg: "3단계" }]);
D.rows(pres, r.slide, r.boxes[0].inner, [["검토 범위", "9월 회수품 12건 전수"], ["판정 기준", "접점 정격 대비 돌입전류"], ["배제 근거", "X-ray 납땜 정상", false], ["확신도", "[추정] 재현시험으로 확정 예정", true]], { lw: 1.2 });
D.steps(pres, r.slide, r.boxes[1].inner, [["재현시험", "정격 1.5배 부하 1만 회"], ["원인 확정", "시험 결과로 가설 판정"], ["대책 적용", "후보 B 적용·횡전개", true]]);

r = S("⑦ 비교표 (넓은 표)", "여러 속성 비교는 1단 넓은 표, 판정 열은 자동 색", "table: colW · hi · judge · num", "1", [{ title: "후보 상세 비교", msg: "숫자 열은 오른쪽 정렬, 강조 행은 연한 빨강" }]);
R = r.boxes[0].inner;
D.table(pres, r.slide, R, ["부품", "정격 전류(A)", "돌입 내량(A)", "단가(원)", "여유율", "신뢰성 시험", "판정"],
  [["현행", "16", "80", "1,200", "0.9배", "완료", "NG"], ["후보 A", "25", "120", "1,850", "1.4배", "필요", "OK"], ["후보 B", "30", "150", "1,600", "1.7배", "필요", "OK"], ["후보 C", "20", "100", "1,300", "1.1배", "필요", "확인 필요"]],
  { colW: [1.6, 1.6, 1.6, 1.6, 1.4, 1.8, R.w - 9.6], hi: 2, judge: 6, num: [1, 2, 3] });

r = S("⑧ 수식 · 도식", "계산은 수식 원래 크기로, 구조는 HTML 도식으로", "equation / figure(kind)", "2-55", [{ title: "계산식", msg: "AC 투입 순간 최대 전류" }, { title: "전류 경로", msg: "NTC가 돌입전류를 제한" }]);
const eq = path.join(figDir, "eq", "E1.png"), dg = path.join(figDir, "diag.png");
if (fs.existsSync(eq)) { [a, b] = D.splitV(r.boxes[0].inner, [1.3, 1]); D.equation(pres, r.slide, a, eq, [917, 259]); D.rows(pres, r.slide, b, [["R_NTC", "NTC 상온 저항"], ["R_line", "배선·기판 저항"]], { lw: 0.9 }); }
else D.T(r.slide, "render_math.py로 eq/E1.png를 먼저 만들 것", { x: r.boxes[0].inner.x, y: r.boxes[0].inner.y, w: 5, h: 0.4, fontSize: 11 });
if (fs.existsSync(dg)) { [a, b] = D.splitV(r.boxes[1].inner, [1, 1]); D.figure(pres, r.slide, a, dg, [1638, 258], "diagram"); D.rows(pres, r.slide, b, [["읽는 법", "왼쪽에서 오른쪽으로 전류 흐름"], ["주의", "회로도는 다시 그리지 않고 원본 사용", true]], { lw: 1.0 }); }
else D.T(r.slide, "render_html.py로 diag.png를 먼저 만들 것", { x: r.boxes[1].inner.x, y: r.boxes[1].inner.y, w: 5, h: 0.4, fontSize: 11 });

D.save(pres, out).then(() => console.log("saved", out));
