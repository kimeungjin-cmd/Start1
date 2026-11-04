"""내용 충실도·시각 다양성 검사 (모델이 바뀌어도 일정 수준 이상 나오게 하는 센서)
사용법: python check_richness.py 결과.pptx [--mode 보고|학습|일반]
본문 장(본문 박스가 있는 장)마다 시각 유형을 찾아 센다. deck.js 부품은 이름표(crp:종류)로, 그 외는 모양으로 판별.
[실패] (종료코드 1)
 - 본문 장이 4장 이상인데 강한 시각 유형(차트·표·숫자카드·타임라인·단계화살표·위치도·사진·수식)이 3종 미만
 - 글 위주 장(카드·행·체크리스트·단계 목록만 있는 장)이 본문의 40% 초과
[경고]
 - 같은 주된 유형이 3장 연속 / 숫자가 하나도 없는 장(보고 모드) / 박스 안 글이 너무 적은 장 / 숫자는 있는데 출처 주석이 없는 장
"""
import sys as _sys
for _s in (_sys.stdout, _sys.stderr):
    try: _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception: pass
import sys, re, argparse
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
EMU = 914400
STRONG = {"a3", "bar", "line", "donut", "pie", "chart", "timeline", "flow", "matrix", "kpi", "table", "picture", "equation", "diagram", "photo", "map", "gantt", "calendar"}
WEAK = {"cards", "rows", "checklist", "steps"}
NAMES = {"a3": "A3 한 장", "bar": "막대", "line": "꺾은선", "donut": "도넛", "pie": "원형", "chart": "차트", "timeline": "타임라인", "flow": "단계 화살표", "matrix": "위치도",
         "kpi": "숫자 카드", "table": "표", "picture": "사진·그림", "equation": "수식", "cards": "카드", "rows": "행 목록", "checklist": "체크리스트", "steps": "단계 목록", "diagram": "도식", "annotate": "주석", "photo": "사진", "map": "지도·위치도", "gantt": "간트", "calendar": "달력"}

def prst(sh):  # 도형 모양 이름(XML 원문) — python-pptx가 모르는 모양에서도 오류 없이
    try:
        v = sh._element.xpath(".//a:prstGeom/@prst"); return v[0] if v else ""
    except Exception: return ""

def kinds(slide):
    ks = []; tagged = set()
    pics = [sh for sh in slide.shapes if sh.shape_type == MSO_SHAPE_TYPE.PICTURE]
    for sh in slide.shapes:
        n = sh.name or ""
        if n.startswith("crp:"): ks.append(n[4:]); tagged.add(n[4:]); continue
        if getattr(sh, "has_chart", False) and sh.has_chart:
            t = str(sh.chart.chart_type).lower()
            ks.append("line" if "line" in t else "donut" if "doughnut" in t else "pie" if "pie" in t else "bar" if ("bar" in t or "column" in t) else "chart")
        elif getattr(sh, "has_table", False) and sh.has_table: ks.append("table")
        elif sh.shape_type == MSO_SHAPE_TYPE.PICTURE and sh.width / EMU > 1.0: ks.append("_pic")
        elif prst(sh) in ("chevron", "homePlate"): ks.append("flow")
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if r.font.size and r.font.size.pt >= 20 and re.search(r"\d", r.text) and sh.top / EMU > 1.9: ks.append("kpi")
    # figure()로 종류를 붙인 그림이 있으면 그 종류로, 없는 그림만 '사진·그림'으로 센다
    if "_pic" in ks:
        ks = [k for k in ks if k != "_pic"] + ([] if tagged & {"flow","matrix","timeline","diagram","photo","map","gantt","calendar","bar","line","donut","equation"} else ["picture"])
    rr = sum(1 for sh in slide.shapes if prst(sh) == "roundRect" and sh.width / EMU > 1.5 and sh.height / EMU > 0.6)
    if rr >= 2: ks.append("cards")
    return ks

def main():
    ap = argparse.ArgumentParser(); ap.add_argument("pptx"); ap.add_argument("--mode", default="보고"); a = ap.parse_args()
    prs = Presentation(a.pptx); body = []; fails = warns = 0; used = set(); prim = []
    for n, sl in enumerate(prs.slides, 1):
        if not any(abs(sh.top / EMU - 1.95) < 0.03 and abs(sh.height / EMU - 5.05) < 0.03 for sh in sl.shapes): continue
        title = next((sh.text_frame.text for sh in sl.shapes if sh.has_text_frame and abs(sh.top / EMU - 0.3) < 0.05), "")
        if title.startswith("유첨"): continue
        ks = kinds(sl); strong = [k for k in ks if k in STRONG]; used |= set(strong)
        body.append((n, ks, strong)); prim.append(strong[0] if strong else (ks[0] if ks else "없음"))
        txt = "\n".join(sh.text_frame.text for sh in sl.shapes if sh.has_text_frame and sh.top / EMU > 1.9 and sh.top / EMU < 7.0)
        nums = re.findall(r"\d", txt); src = any(sh.has_text_frame and sh.text_frame.text.strip().startswith("※") and sh.top / EMU > 7.0 for sh in sl.shapes)
        label = ", ".join(sorted({NAMES.get(k, k) for k in ks})) or "없음"
        print(f"[{n}장] 시각 유형: {label}")
        if not strong: print(f"   - 경고: 글 위주 장 (강한 시각 유형 없음)"); warns += 1
        if a.mode == "보고" and not nums: print(f"   - 경고: 숫자 근거가 하나도 없음"); warns += 1
        if len(txt.replace(" ", "")) < 80 and not any(k in ks for k in ("bar", "line", "donut", "picture", "matrix", "table")): print(f"   - 경고: 박스 안 내용이 너무 적음"); warns += 1
        if nums and not src: print(f"   - 경고: 숫자가 있는데 하단 출처 주석(※)이 없음"); warns += 1
    nb = len(body)
    weak_n = sum(1 for _, _, st in body if not st)
    for i in range(2, len(prim)):
        if prim[i] == prim[i - 1] == prim[i - 2] and prim[i] != "없음":
            print(f"[경고] {body[i-2][0]}~{body[i][0]}장 주된 시각 유형이 3장 연속 '{NAMES.get(prim[i], prim[i])}' — 내용에 맞는 다른 유형 검토"); warns += 1
    need = 3 if nb >= 4 else (2 if nb >= 2 else 1)
    print(f"\n본문 {nb}장 / 강한 시각 유형 {len(used)}종: {', '.join(NAMES.get(k, k) for k in sorted(used)) or '없음'}")
    if nb >= 2 and len(used) < need:
        print(f"[실패] 시각 유형이 {need}종 미만 — layouts-and-boxes.md 패턴표와 deck.js 시각화 부품(chartBar·chartLine·chartDonut·timeline·flow·matrix·kpis·table)에서 내용에 맞는 유형을 골라 다양화할 것"); fails += 1
    if nb >= 4 and weak_n / nb > 0.4:
        print(f"[실패] 글 위주 장이 {weak_n}/{nb}장 — 장마다 시각 요소를 넣을 것"); fails += 1
    custom = sum(1 for _, ks, _ in body if any(k in ks for k in ("annotate", "diagram", "map", "gantt", "calendar", "photo")))
    if nb >= 5 and custom == 0:
        print("[경고] 맞춤 시각화 장이 0개 — 부품만으로 만들면 문서가 단조롭다. 핵심 근거 장 1~2개는 주석(annotate)을 단 차트나 HTML 도식(figure)으로 내용에 맞게 설계할 것"); warns += 1
    print(f"충실도 이슈: 실패 {fails}건, 경고 {warns}건"); sys.exit(1 if fails else 0)

if __name__ == "__main__":
    main()
