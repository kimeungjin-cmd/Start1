// 회사 양식 PPT 생성 라이브러리
// 고정 영역(헤더 4요소, 박스 틀, 레이아웃 좌표)은 이 파일이 그린다. AI는 boxes[].inner 안쪽만 그린다.

// ▼▼ 회사 실제 값으로 교체할 곳 (색·폰트) ▼▼
const THEME = {
  font: "맑은 고딕",
  main: "1F3864",      // 소제목 글자·가로선·세로막대·박스 메시지 글자. 사내 보고서 남색과 유사해 확정 (삼성 BI #1428A0은 너무 쨍해서 미사용)
  band: "1F3864",      // 본문 박스 소제목 띠 배경 (회사 색이 글자색과 다르면 여기만 바꿈)
  dark: "222222", gray: "6B6B6B", muted: "A6A6A6",
  boxLine: "595959",   // 본문 박스 외곽선
  tint: "EEF2F8", mid: "8EA3C4", light: "DCE3EE",
  red: "C00000", redTint: "FBE9E9",
};
// ▲▲ 회사 실제 값으로 교체할 곳 ▲▲

const pptxgen = require("pptxgenjs");
const X0 = 0.5, W = 12.33, GAP = 0.25, BY = 1.95, BH = 5.05, TITLE_H = 0.42;

// 레이아웃 5종. 박스 좌표는 여기서만 계산한다.
const LAYOUTS = { "1": [1], "2-55": [5, 5], "2-64": [6, 4], "2-73": [7, 3], "3": [1, 1, 1] };
// 글자 예산(넘침 사전 차단). 초과 시 경고 → 문장을 줄일 것
const BUDGET = { title: 28, message: 45, sub: 60, boxTitle: 24 };

function warnLen(kind, text) {
  const n = [...String(text)].length;
  if (BUDGET[kind] && n > BUDGET[kind]) console.warn(`[글자 예산 초과] ${kind} ${n}/${BUDGET[kind]}자: ${text}`);
}
function createDeck() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.theme = { headFontFace: THEME.font, bodyFontFace: THEME.font };
  return pres;
}
function T(s, text, o) {
  s.addText(text, Object.assign({ fontFace: THEME.font, margin: 0, isTextBox: true }, o));
}
// 고정: 소제목 / 가로선 / 세로막대 + 핵심 메시지 / 부연설명 / (선택) 하단 출처
function header(pres, s, { title, message, sub, source }) {
  warnLen("title", title); warnLen("message", message); if (sub) warnLen("sub", sub);
  T(s, title, { x: X0, y: 0.3, w: W, h: 0.5, fontSize: 22, bold: true, color: THEME.main });
  s.addShape(pres.shapes.LINE, { x: X0, y: 0.88, w: W, h: 0, line: { color: THEME.main, width: 1.5 } });
  s.addShape(pres.shapes.RECTANGLE, { x: X0, y: 1.06, w: 0.07, h: 0.33, fill: { color: THEME.main }, line: { color: THEME.main } });
  T(s, message, { x: 0.7, y: 1.0, w: 12.13, h: 0.45, fontSize: 18, bold: true, color: THEME.dark });
  if (sub) T(s, sub, { x: 0.7, y: 1.45, w: 12.13, h: 0.32, fontSize: 12, color: THEME.gray });
  if (source) T(s, source, { x: X0, y: 7.1, w: W - 0.8, h: 0.25, fontSize: 9, color: THEME.muted });
  // 페이지 번호 "현재 / 전체": 현재 번호는 슬라이드 번호 필드(순서 바꿔도 자동 갱신), 전체 수는 save() 때 기입
  s.slideNumber = { x: X0 + W - 0.84, y: 7.1, w: 0.5, h: 0.25, fontFace: THEME.font, fontSize: 9, color: THEME.gray, align: "right" };
  (pres._crpSlides = pres._crpSlides || []).push(s);
}

// ═════════ 재사용 부품 (박스 inner 영역 A={x,y,w,h} 안에 배치) ═════════
// 직접 좌표를 짜기 전에 먼저 이 부품으로 해결한다. 높이는 모두 A.h에서 나눠 계산하므로 박스를 넘치지 않는다.
// 카드 여러 장 세로 배치: items=[{title, body, tag?, big?, red?}]
function cards(pres, s, A, items, { gap = 0.12 } = {}) {
  _tag(s, "cards", A);
  const h = (A.h - gap * (items.length - 1)) / items.length;
  items.forEach((it, i) => {
    const y = A.y + i * (h + gap), fg = it.red ? THEME.red : THEME.main, bg = it.red ? THEME.redTint : THEME.tint;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: A.x, y, w: A.w, h, rectRadius: 0.08, fill: { color: bg }, line: { color: bg } });
    let tx = A.x + 0.18, top = y + 0.08;
    if (it.tag) { pill(pres, s, tx, y + 0.12, 1.0, 0.28, it.tag, "FFFFFF", fg); tx += 1.1; }
    T(s, it.title, { x: tx, y: top, w: A.x + A.w - 0.18 - tx, h: 0.36, fontSize: 13, bold: true, color: it.tag ? THEME.dark : fg, valign: "middle" });
    let by = y + 0.46;
    if (it.big) { T(s, it.big, { x: A.x + 0.18, y: by, w: A.w - 0.36, h: 0.36, fontSize: 16, bold: true, color: fg, valign: "middle" }); by += 0.38; }
    T(s, it.body || "", { x: A.x + 0.18, y: by, w: A.w - 0.36, h: Math.max(y + h - by - 0.06, 0.25), fontSize: 11, color: THEME.dark, valign: "top" });
  });
}
// 라벨-설명 행 목록: items=[[라벨, 설명, red?]]
function rows(pres, s, A, items, { lw = 1.4 } = {}) {
  _tag(s, "rows", A);
  const rh = A.h / items.length;
  items.forEach(([k, v, red], i) => {
    const y = A.y + i * rh;
    T(s, k, { x: A.x, y, w: lw, h: rh - 0.08, fontSize: 12, bold: true, color: red ? THEME.red : THEME.main, valign: "middle" });
    T(s, v, { x: A.x + lw + 0.1, y, w: A.w - lw - 0.1, h: rh - 0.08, fontSize: 11, color: THEME.dark, valign: "middle" });
    if (i < items.length - 1) s.addShape(pres.shapes.LINE, { x: A.x, y: y + rh - 0.04, w: A.w, h: 0, line: { color: "E3E3E3", width: 0.5 } });
  });
}
// 큰 숫자 카드 가로 배치: items=[{num, unit?, label, red?}]
function kpis(pres, s, A, items, { gap = 0.15 } = {}) {
  _tag(s, "kpi", A);
  const w = (A.w - gap * (items.length - 1)) / items.length;
  items.forEach((it, i) => {
    const x = A.x + i * (w + gap), fg = it.red ? THEME.red : THEME.main, bg = it.red ? THEME.redTint : THEME.tint;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: A.y, w, h: A.h, rectRadius: 0.08, fill: { color: bg }, line: { color: bg } });
    s.addText([{ text: String(it.num), options: { fontSize: 26, bold: true, color: fg } }, { text: it.unit ? " " + it.unit : "", options: { fontSize: 13, bold: true, color: THEME.gray } }],
      { x, y: A.y + A.h * 0.12, w, h: A.h * 0.5, fontFace: THEME.font, align: "center", valign: "middle", margin: 0 });
    T(s, it.label, { x: x + 0.1, y: A.y + A.h * 0.62, w: w - 0.2, h: A.h * 0.34, fontSize: 11, color: THEME.dark, align: "center", valign: "top" });
  });
}
// 번호 단계 목록: items=[[제목, 설명, red?]]
function steps(pres, s, A, items) {
  _tag(s, "steps", A);
  const rh = A.h / items.length;
  items.forEach(([t, d, red], i) => {
    const y = A.y + i * rh, col = red ? THEME.red : THEME.main;
    s.addShape(pres.shapes.OVAL, { x: A.x, y: y + (rh - 0.42) / 2, w: 0.42, h: 0.42, fill: { color: col }, line: { color: col } });
    T(s, String(i + 1), { x: A.x, y: y + (rh - 0.42) / 2, w: 0.42, h: 0.42, fontSize: 13, bold: true, color: "FFFFFF", align: "center", valign: "middle" });
    T(s, t, { x: A.x + 0.6, y: y + rh / 2 - 0.36, w: A.w - 0.6, h: 0.34, fontSize: 13, bold: true, color: red ? THEME.red : THEME.dark, valign: "middle" });
    T(s, d, { x: A.x + 0.6, y: y + rh / 2 - 0.02, w: A.w - 0.6, h: 0.34, fontSize: 11, color: THEME.gray, valign: "top" });
  });
}
// 체크리스트: items=[[항목, 설명, "필수"|"선택"]]
function checklist(pres, s, A, items) {
  _tag(s, "checklist", A);
  const rh = A.h / items.length;
  items.forEach(([t, d, tag], i) => {
    const y = A.y + i * rh, must = tag === "필수";
    s.addShape(pres.shapes.RECTANGLE, { x: A.x, y: y + 0.08, w: 0.3, h: 0.3, fill: { color: "FFFFFF" }, line: { color: THEME.main, width: 1.5 } });
    T(s, t, { x: A.x + 0.5, y, w: A.w - 1.8, h: 0.4, fontSize: 13, bold: true, color: THEME.dark, valign: "middle" });
    T(s, d, { x: A.x + 0.5, y: y + 0.38, w: A.w - 0.5, h: 0.3, fontSize: 10.5, color: THEME.gray, valign: "middle" });
    if (tag) pill(pres, s, A.x + A.w - 1.1, y + 0.08, 0.95, 0.28, tag, must ? "FFFFFF" : THEME.main, must ? THEME.main : THEME.tint);
  });
}
// 표: head=[...], body=[[...]], 강조 행 번호 hi(선택). 행 높이 합이 A.h를 넘으면 오류
function table(pres, s, A, head, body, { colW, hi = -1, rowH = 0.45, fontSize = 11, judge = -1, num = [] } = {}) {
  // hi: 강조 행 번호(연한 빨강 배경) / judge: 판정 열 번호 — OK·적합은 메인색, NG·부적합·초과는 빨강, 확인 필요는 회색 / num: 숫자 열(오른쪽 정렬)
  _tag(s, "table", A);
  if ((body.length + 1) * rowH > A.h + 0.01) throw new Error(`표가 박스를 넘음: ${(body.length + 1) * rowH}in > ${A.h.toFixed(2)}in — 행을 줄이거나 유첨으로`);
  const H = head.map(t => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: THEME.band }, align: "center" } }));
  const jc = v => /NG|부적합|초과|미달|불량/.test(v) ? THEME.red : /OK|적합|양호|통과/.test(v) ? THEME.main : THEME.gray;
  const B = body.map((r, i) => r.map((c, j) => {
    const v = String(c), o = { bold: j === 0, color: j === 0 ? THEME.main : THEME.dark, align: j === 0 ? "left" : num.includes(j) ? "right" : "center", fill: { color: i === hi ? THEME.redTint : (i % 2 ? "F7F9FC" : "FFFFFF") } };
    if (j === judge) Object.assign(o, { bold: true, color: jc(v) });
    return { text: v, options: o };
  }));
  s.addTable([H, ...B], { x: A.x, y: A.y, w: A.w, colW, fontFace: THEME.font, fontSize, rowH, border: { type: "solid", pt: 0.5, color: "D9D9D9" }, valign: "middle", margin: 0.06 });
  return A.y + (body.length + 1) * rowH;  // 표 아래 다음 요소 시작 y
}
// 영역 나누기 도우미: 세로로 [비율...] 분할 → [{x,y,w,h}]
function splitV(A, ratios, gap = 0.15) {
  const tot = ratios.reduce((a, b) => a + b, 0), avail = A.h - gap * (ratios.length - 1); let y = A.y;
  return ratios.map(r => { const h = avail * r / tot, o = { x: A.x, y, w: A.w, h }; y += h + gap; return o; });
}
function splitH(A, ratios, gap = 0.2) {
  const tot = ratios.reduce((a, b) => a + b, 0), avail = A.w - gap * (ratios.length - 1); let x = A.x;
  return ratios.map(r => { const w = avail * r / tot, o = { x, y: A.y, w, h: A.h }; x += w + gap; return o; });
}

// 수식 이미지(render_math.py 결과)를 '원래 크기'로 넣기 — 박스에 맞춰 늘리지 않음(늘리면 글자가 본문보다 커짐)
// sizePx: render_math.py가 출력한 [w,h], scale: 렌더링 배율(기본 3), align: "left"|"center"
function equation(pres, s, A, path, sizePx, { scale = 3, align = "center", maxScale = 1 } = {}) {
  _tag(s, "equation", A);
  let w = sizePx[0] / scale / 96, h = sizePx[1] / scale / 96;
  const k = Math.min(maxScale, A.w / w, A.h / h); w *= k; h *= k;
  const x = align === "left" ? A.x : A.x + (A.w - w) / 2, y = A.y + (A.h - h) / 2;
  s.addImage({ path, x, y, w, h });
  return { x, y, w, h };
}


// ═════════ 시각화 부품 (모델이 좌표를 짜지 않고 고르기만 하면 되게) ═════════
// 모든 시각화 부품은 첫 도형에 objectName "crp:<종류>"를 붙인다 → check_richness.py가 장별 시각 유형을 센다.
function _tag(s, kind, A) { s.addShape("rect", { x: A.x, y: A.y, w: 0.01, h: 0.01, fill: { color: "FFFFFF", transparency: 100 }, line: { color: "FFFFFF", transparency: 100 }, objectName: "crp:" + kind }); }
const _mix = (hex, w) => [0, 2, 4].map(i => parseInt(hex.substr(i, 2), 16)).map(v => Math.round(v + (255 - v) * w).toString(16).padStart(2, "0")).join("").toUpperCase();

// 가로 막대: 강조 1개(메인색) + 관련(연한 메인색) + 나머지 회색. 큰 값 순으로 넣을 것
function chartBar(pres, s, A, { labels, values, unit = "", highlight = [0], related = [], max } = {}) {
  _tag(s, "bar", A);
  const top = max || Math.max(...values) * 1.2, P = { x: 0.28, y: 0.04, w: 0.68, h: 0.92 };
  const colors = labels.map((_, i) => highlight.includes(i) ? THEME.main : related.includes(i) ? THEME.mid : THEME.muted);
  s.addChart(pres.charts.BAR, [{ name: "값", labels, values }], { x: A.x, y: A.y, w: A.w, h: A.h, barDir: "bar", catAxisOrientation: "maxMin", layout: P,
    chartColors: colors, barGapWidthPct: 45, showValue: true, dataLabelFontSize: 11, dataLabelFontFace: THEME.font, dataLabelColor: THEME.dark,
    dataLabelFormatCode: unit ? `0.##"${unit}"` : "0.##", catAxisLabelFontFace: THEME.font, catAxisLabelFontSize: 11, catAxisLabelColor: THEME.dark,
    catAxisLineShow: false, valAxisHidden: true, valGridLine: { style: "none" }, valAxisMinVal: 0, valAxisMaxVal: top, showLegend: false });
  const n = labels.length;
  // 반환: barEnd(i) → i번째 막대 끝 {px, py} (annotate용)
  return { barEnd: i => ({ px: P.x + P.w * values[i] / top - 0.02, py: P.y + P.h * (i + 0.5) / n }) };  // 값 라벨을 가리지 않게 막대 끝 살짝 안쪽
}
// 꺾은선: 추이. series=[{name, values}], 첫 계열이 메인색. 기준선(target) 선택
function chartLine(pres, s, A, { labels, series, unit = "", target, min = 0, max } = {}) {
  _tag(s, "line", A);
  const all = series.flatMap(x => x.values).concat(target != null ? [target] : []);
  const top = max != null ? max : Math.max(...all) * 1.25;
  const data = series.map(x => ({ name: x.name, labels, values: x.values }));
  if (target != null) data.push({ name: `기준 ${target}${unit}`, labels, values: labels.map(() => target) });
  const P = { x: 0.04, y: 0.06, w: 0.92, h: 0.72 };   // 그림 영역을 고정해 두어 점 위치를 계산할 수 있게
  s.addChart(pres.charts.LINE, data, { x: A.x, y: A.y, w: A.w, h: A.h, layout: P, chartColors: [THEME.main, THEME.mid, THEME.muted, THEME.red].slice(0, series.length).concat(target != null ? [THEME.red] : []),
    lineSize: 2.5, lineDataSymbol: "circle", lineDataSymbolSize: 7, showValue: target == null && series.length === 1, dataLabelFontSize: 10, dataLabelPosition: "t", dataLabelFormatCode: unit ? `0.##"${unit}"` : "0.##",
    catAxisLabelFontFace: THEME.font, catAxisLabelFontSize: 11, valAxisHidden: true, valAxisMinVal: min, valAxisMaxVal: top, valGridLine: { style: "none" },
    showLegend: series.length > 1 || target != null, legendPos: "b", legendFontSize: 10, legendFontFace: THEME.font });
  const n = labels.length;
  // 반환: point(i, v) → annotate에 넣을 {px, py} (영역 A 기준 비율)
  return { point: (i, v) => ({ px: P.x + P.w * (i + 0.5) / n, py: P.y + P.h * (1 - (v - min) / (top - min)) }) };
}
// 도넛 + 가운데 큰 숫자 + 오른쪽 순위 범례. 항목 6개 이하
function chartDonut(pres, s, A, { labels, values, center, centerSub = "", unit = "" } = {}) {
  _tag(s, "donut", A);
  const d = Math.min(A.h, A.w * 0.48), cols = labels.map((_, i) => [THEME.main, THEME.mid, THEME.light, THEME.muted, "D9D9D9", "EEEEEE"][i] || "EEEEEE");
  s.addChart(pres.charts.DOUGHNUT, [{ name: "비중", labels, values }], { x: A.x, y: A.y + (A.h - d) / 2, w: d, h: d, holeSize: 62, chartColors: cols, showLegend: false, showValue: false, showPercent: false, dataBorder: { pt: 1.5, color: "FFFFFF" } });
  s.addText([{ text: String(center), options: { fontSize: 24, bold: true, color: THEME.main, breakLine: true } }, { text: centerSub, options: { fontSize: 10, color: THEME.gray } }],
    { x: A.x, y: A.y + A.h / 2 - 0.45, w: d, h: 0.9, fontFace: THEME.font, align: "center", valign: "middle", margin: 0 });
  const lx = A.x + d + 0.25, lw = A.w - d - 0.25, rh = Math.min(0.5, A.h / labels.length);
  const y0 = A.y + (A.h - rh * labels.length) / 2;
  labels.forEach((l, i) => {
    const y = y0 + i * rh;
    s.addShape(pres.shapes.RECTANGLE, { x: lx, y: y + rh / 2 - 0.1, w: 0.2, h: 0.2, fill: { color: cols[i] }, line: { color: cols[i] } });
    T(s, l, { x: lx + 0.3, y, w: lw - 1.1, h: rh, fontSize: 11, bold: i === 0, color: i === 0 ? THEME.main : THEME.dark, valign: "middle" });
    T(s, values[i] + unit, { x: lx + lw - 0.8, y, w: 0.8, h: rh, fontSize: i === 0 ? 13 : 11, bold: true, color: i === 0 ? THEME.main : THEME.gray, align: "right", valign: "middle" });
  });
}
// 시간 흐름: items=[{date, title, desc, red}] 위아래 번갈아 배치
function timeline(pres, s, A, items) {
  _tag(s, "timeline", A);
  const ly = A.y + A.h / 2, x0 = A.x + 0.5, x1 = A.x + A.w - 0.5, step = items.length > 1 ? (x1 - x0) / (items.length - 1) : 0, bw = Math.min(1.9, Math.max(step * 1.6, 1.2));
  s.addShape(pres.shapes.LINE, { x: x0, y: ly, w: x1 - x0, h: 0, line: { color: THEME.mid, width: 3 } });
  items.forEach((it, i) => {
    const cx = x0 + i * step, up = i % 2 === 0, col = it.red ? THEME.red : THEME.main, bx = Math.min(Math.max(cx - bw / 2, A.x), A.x + A.w - bw);
    s.addShape(pres.shapes.OVAL, { x: cx - 0.12, y: ly - 0.12, w: 0.24, h: 0.24, fill: { color: col }, line: { color: "FFFFFF", width: 1.5 } });
    const by = up ? ly - Math.min(1.5, A.h / 2 - 0.05) : ly + 0.3, bh = Math.min(1.5, A.h / 2 - 0.3);
    T(s, it.date, { x: bx, y: by, w: bw, h: 0.28, fontSize: 11.5, bold: true, color: col, align: "center" });
    T(s, it.title, { x: bx, y: by + 0.28, w: bw, h: 0.3, fontSize: 12, bold: true, color: THEME.dark, align: "center" });
    if (it.desc) T(s, it.desc, { x: bx, y: by + 0.58, w: bw, h: Math.max(bh - 0.6, 0.3), fontSize: 10, color: THEME.gray, align: "center", valign: "top" });
  });
}
// 단계 화살표: items=[{title, desc, state}] state: done(회색) now(메인) plan(연한) risk(빨강)
function flow(pres, s, A, items) {
  _tag(s, "flow", A);
  const n = items.length, ov = 0.15, cw = (A.w + ov * (n - 1)) / n, chH = Math.min(0.62, A.h * 0.3);
  const st = { done: [THEME.muted, "FFFFFF"], now: [THEME.main, "FFFFFF"], plan: [THEME.light, THEME.main], risk: [THEME.red, "FFFFFF"] };
  items.forEach((it, i) => {
    const x = A.x + i * (cw - ov), [bg, fg] = st[it.state || "now"];
    s.addShape(pres.shapes.CHEVRON, { x, y: A.y, w: cw, h: chH, fill: { color: bg }, line: { color: "FFFFFF", width: 1.5 } });
    T(s, it.title, { x: x + 0.15, y: A.y, w: cw - 0.35, h: chH, fontSize: 11.5, bold: true, color: fg, align: "center", valign: "middle" });
    if (it.desc) T(s, it.desc, { x: x + 0.05, y: A.y + chH + 0.12, w: cw - 0.25, h: A.h - chH - 0.15, fontSize: 10.5, color: THEME.dark, align: "center", valign: "top" });
  });
}
// 2×2 위치도: items=[{name, x, y, note, red}] x·y는 0~1. goodZone: "tl"|"tr"|"bl"|"br" 음영
function matrix(pres, s, A, { xLabel, yLabel, items, goodZone = "tl", goodText = "" } = {}) {
  _tag(s, "matrix", A);
  const gx = A.x + 0.15, gy = A.y + 0.35, gw = A.w - 0.25, gh = A.h - 0.75;
  const zx = goodZone[1] === "l" ? gx : gx + gw / 2, zy = goodZone[0] === "t" ? gy : gy + gh / 2;
  s.addShape(pres.shapes.RECTANGLE, { x: zx, y: zy, w: gw / 2, h: gh / 2, fill: { color: THEME.tint }, line: { color: THEME.tint } });
  if (goodText) T(s, goodText, { x: zx + 0.1, y: zy + 0.05, w: gw / 2 - 0.2, h: 0.3, fontSize: 10.5, bold: true, color: THEME.main });
  s.addShape(pres.shapes.LINE, { x: gx, y: gy + gh, w: gw, h: 0, line: { color: THEME.muted, width: 1.25 } });
  s.addShape(pres.shapes.LINE, { x: gx, y: gy, w: 0, h: gh, line: { color: THEME.muted, width: 1.25 } });
  T(s, "↑ " + yLabel, { x: gx, y: A.y, w: gw / 2, h: 0.3, fontSize: 10.5, bold: true, color: THEME.gray });
  T(s, xLabel + " →", { x: gx, y: gy + gh + 0.08, w: gw, h: 0.3, fontSize: 10.5, bold: true, color: THEME.gray, align: "right" });
  items.forEach(it => {
    const cx = gx + it.x * gw, cy = gy + (1 - it.y) * gh, col = it.red ? THEME.red : THEME.main, lw = 1.7;
    const right = cx + 0.16 + lw <= A.x + A.w;   // 오른쪽에 자리가 있으면 오른쪽, 없으면 왼쪽에 이름표
    s.addShape(pres.shapes.OVAL, { x: cx - 0.11, y: cy - 0.11, w: 0.22, h: 0.22, fill: { color: col }, line: { color: "FFFFFF", width: 1.5 } });
    T(s, it.name + (it.note ? "\n" + it.note : ""), { x: right ? cx + 0.16 : cx - 0.16 - lw, y: Math.max(cy - 0.2, A.y), w: lw, h: 0.5, fontSize: 10.5, bold: true, color: col, align: right ? "left" : "right", valign: "top" });
  });
}

// 주석(콜아웃): 차트·그림 위 특정 지점에 "여기가 핵심" 표시. px·py는 영역 A 안의 비율(0~1), 상자는 dx·dy(인치)만큼 떨어뜨림
// 차트 종류보다 '어디를 봐야 하는지' 짚어 주는 주석이 시각 자료를 단순하지 않게 만든다. 렌더링으로 위치를 꼭 확인할 것
function annotate(pres, s, A, { px, py, text, dx = 0.6, dy = -0.6, w = 1.9, red = true } = {}) {
  _tag(s, "annotate", A);
  const col = red ? THEME.red : THEME.main, x = A.x + px * A.w, y = A.y + py * A.h;
  let bx = x + dx, by = y + dy; const h = 0.5;
  bx = Math.min(Math.max(bx, A.x), A.x + A.w - w); by = Math.min(Math.max(by, A.y), A.y + A.h - h);
  const ex = bx + (bx > x ? 0 : w), ey = by + h / 2;
  s.addShape(pres.shapes.LINE, { x: Math.min(x, ex), y: Math.min(y, ey), w: Math.abs(ex - x) || 0.001, h: Math.abs(ey - y) || 0.001, flipV: (ey < y) !== (ex < x), line: { color: col, width: 1.25, dashType: "dash" } });
  s.addShape(pres.shapes.OVAL, { x: x - 0.09, y: y - 0.09, w: 0.18, h: 0.18, fill: { color: "FFFFFF" }, line: { color: col, width: 2 } });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: bx, y: by, w, h, rectRadius: 0.06, fill: { color: "FFFFFF" }, line: { color: col, width: 1.25 } });
  T(s, text, { x: bx + 0.08, y: by, w: w - 0.16, h, fontSize: 10, bold: true, color: col, valign: "middle" });
}

// HTML로 그린 그림(render_html.py 결과)·사진 넣기 + 시각 유형 표시. kind: flow|matrix|timeline|bar|line|donut|diagram|photo|map 등
function figure(pres, s, A, path, sizePx, kind = "diagram", caption = null) {
  _tag(s, kind, A);
  return photo(pres, s, A, path, sizePx, caption);
}

// --fit 으로 영역 크기 그대로 렌더링한 HTML 그림을 '늘리지 않고' 1:1로 넣기 (글자 크기가 본문과 일치)
function figureFit(pres, s, A, path, kind = "diagram") {
  _tag(s, kind, A);
  s.addImage({ path, x: A.x, y: A.y, w: A.w, h: A.h });
  return A;
}

// 표지 (고정 영역): 제목 / 부제 / 작성 조직 / 날짜. 페이지 번호는 넣지 않음
function cover(pres, { title, subtitle, org, date }) {
  const s = pres.addSlide();
  pres._crpCovers = (pres._crpCovers || 0) + 1; // 표지도 전체 장 수에 포함(번호 표시는 안 함)
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.35, h: 7.5, fill: { color: THEME.main }, line: { color: THEME.main } });
  T(s, title, { x: 1.2, y: 2.3, w: 11.0, h: 1.0, fontSize: 36, bold: true, color: THEME.main, valign: "bottom" });
  s.addShape(pres.shapes.LINE, { x: 1.2, y: 3.45, w: 11.0, h: 0, line: { color: THEME.main, width: 1.5 } });
  if (subtitle) T(s, subtitle, { x: 1.2, y: 3.6, w: 11.0, h: 0.5, fontSize: 16, color: THEME.gray });
  const meta = [org, date].filter(Boolean).join("\n");
  if (meta) T(s, meta, { x: 1.2, y: 5.6, w: 11.0, h: 0.8, fontSize: 14, bold: true, color: THEME.dark, valign: "bottom", paraSpaceAfter: 2 });
  return s;
}
// 저장: 모든 장에 "/ 전체" 기입 후 파일 쓰기. 반드시 pres.writeFile 대신 이것을 쓴다
// opts.reserveCover: 회사 표지를 나중에 손으로 맨 앞에 붙일 예정이면 true (전체 수에 1장 미리 포함)
function save(pres, fileName, opts = {}) {
  const all = pres._crpSlides || [], n = all.length + (pres._crpCovers || 0) + (opts.reserveCover ? 1 : 0);
  all.forEach(s => T(s, `\u00A0/ ${n}`, { x: X0 + W - 0.43, y: 7.1, w: 0.43, h: 0.25, fontSize: 9, color: THEME.gray, align: "left" }));
  return pres.writeFile({ fileName });
}
// 고정: 본문 박스 틀. msg가 있으면 박스 내부 첫 줄(박스 메시지)로 그림
function box(pres, s, x, w, title, msg) {
  warnLen("boxTitle", title);
  s.addShape(pres.shapes.RECTANGLE, { x, y: BY, w, h: BH, fill: { color: "FFFFFF" }, line: { color: THEME.boxLine, width: 1 } });
  s.addShape(pres.shapes.RECTANGLE, { x, y: BY, w, h: TITLE_H, fill: { color: THEME.band }, line: { color: THEME.band, width: 0.75 } });
  T(s, title, { x: x + 0.15, y: BY, w: w - 0.3, h: TITLE_H, fontSize: 14, bold: true, color: "FFFFFF", valign: "middle" });
  let top = BY + TITLE_H + 0.13;
  if (msg) {
    T(s, msg, { x: x + 0.2, y: top, w: w - 0.4, h: 0.35, fontSize: 13, bold: true, color: THEME.main });
    top += 0.45;
  }
  return { x, y: BY, w, h: BH, inner: { x: x + 0.2, y: top, w: w - 0.4, h: BY + BH - top - 0.15 } };
}
// 한 장 생성 → { slide, boxes:[{inner:{x,y,w,h}}] }
function slide(pres, { title, message, sub, source, layout = "2-55", boxes = [], notes }) {
  const ratio = LAYOUTS[layout];
  if (!ratio) throw new Error(`알 수 없는 레이아웃: ${layout} (가능: ${Object.keys(LAYOUTS).join(", ")})`);
  if (boxes.length !== ratio.length) throw new Error(`레이아웃 ${layout}은 박스 ${ratio.length}개 필요 (받은 개수 ${boxes.length})`);
  const s = pres.addSlide();
  header(pres, s, { title, message, sub, source });
  const avail = W - GAP * (ratio.length - 1), sum = ratio.reduce((a, b) => a + b, 0);
  let x = X0; const out = [];
  ratio.forEach((r, i) => { const w = avail * r / sum; out.push(box(pres, s, x, w, boxes[i].title, boxes[i].msg)); x += w + GAP; });
  if (notes) s.addNotes(notes);
  return { slide: s, boxes: out };
}
// ── 보조 도구 ──
function pill(pres, s, x, y, w, h, text, fg, bg) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: h / 2, fill: { color: bg }, line: { color: bg } });
  T(s, text, { x, y, w, h, fontSize: 10, bold: true, color: fg, align: "center", valign: "middle" });
}
// 박스 하단 주석(기준·출처·"상세: 유첨 N")
function boxNote(s, inner, text) {
  T(s, text, { x: inner.x, y: inner.y + inner.h - 0.25, w: inner.w, h: 0.25, fontSize: 10, color: THEME.gray });
}
// 사진: 비율 유지로 영역 안에 맞춤 + 캡션. 사진 위 표시(빨간 원 등)는 별도 도형으로 얹을 것
function photo(pres, s, area, path, sizePx, caption) {
  const capH = caption ? 0.3 : 0;
  const ar = sizePx[0] / sizePx[1], aw = area.w, ah = area.h - capH;
  let w = aw, h = aw / ar; if (h > ah) { h = ah; w = ah * ar; }
  const x = area.x + (aw - w) / 2, y = area.y;
  s.addImage({ path, x, y, w, h });
  if (caption) T(s, caption, { x: area.x, y: y + h + 0.05, w: aw, h: 0.25, fontSize: 10, color: THEME.gray, align: "center" });
  return { x, y, w, h };
}
module.exports = { figureFit, annotate, figure, chartBar, chartLine, chartDonut, timeline, flow, matrix, equation, cards, rows, kpis, steps, checklist, table, splitV, splitH, save, cover, THEME, LAYOUTS, BUDGET, createDeck, slide, header, box, pill, boxNote, photo, T };
