"""이미 만든 PPT에서 '특정 장의 특정 박스' 내용만 HTML 그림으로 바꾸기 (선택 업그레이드)
PowerPoint에서 손으로 고친 파일에도 쓸 수 있다(생성 스크립트를 다시 돌리지 않음).

1) 크기 확인:  python upgrade_box.py 결과.pptx --slide 3 --box 1 --info
   → 박스 안쪽 영역 크기(인치)와 지금 들어 있는 내용(글·차트 데이터)을 출력. 이 내용으로 HTML을 그린다
2) HTML 그리기: python render_html.py 그림.html 그림.png --fit W H   (1의 크기 그대로)
3) 교체:      python upgrade_box.py 결과.pptx --slide 3 --box 1 --image 그림.png --kind bar [--out 새파일.pptx]
   → 박스 제목 띠·박스 메시지는 그대로 두고 안쪽 내용만 그림으로 교체.
     원래 내용(글·차트 데이터)은 슬라이드 노트에 '[업그레이드 전 원본]'으로 보존(누락 검사·되돌리기용)
박스 번호는 왼쪽부터 1, 2, 3.
"""
import sys as _sys
for _s in (_sys.stdout, _sys.stderr):
    try: _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception: pass
import argparse, sys
from pptx import Presentation
from pptx.util import Emu
EMU = 914400; BY, BH, TITLE_H = 1.95, 5.05, 0.42

def boxes_of(slide):
    bs = [sh for sh in slide.shapes if abs(sh.top / EMU - BY) < 0.03 and abs(sh.height / EMU - BH) < 0.03]
    return sorted(bs, key=lambda s: s.left)

def inner_of(slide, box):
    x, w = box.left / EMU, box.width / EMU
    top = BY + TITLE_H + 0.13
    msg = [sh for sh in slide.shapes if sh.has_text_frame and abs(sh.top / EMU - top) < 0.03 and abs(sh.left / EMU - (x + 0.2)) < 0.03]
    if msg: top += 0.45
    return (x + 0.2, top, w - 0.4, BY + BH - top - 0.15), msg

def inside(sh, A):
    x, y, w, h = A; sx, sy = sh.left / EMU, sh.top / EMU; sw, shh = (sh.width or 0) / EMU, (sh.height or 0) / EMU
    return sx >= x - 0.05 and sy >= y - 0.05 and sx + sw <= x + w + 0.05 and sy + shh <= y + h + 0.05

def describe(sh):
    if getattr(sh, "has_chart", False) and sh.has_chart:
        out = []
        for p in sh.chart.plots:
            cats = list(p.categories)
            for se in p.series: out.append(f"차트[{se.name}] " + ", ".join(f"{c}={v}" for c, v in zip(cats, se.values)))
        return "\n".join(out)
    if getattr(sh, "has_table", False) and sh.has_table:
        return "표\n" + "\n".join(" | ".join(c.text for c in r.cells) for r in sh.table.rows)
    if sh.has_text_frame and sh.text_frame.text.strip(): return sh.text_frame.text.strip()
    return ""

def main():
    ap = argparse.ArgumentParser(); ap.add_argument("pptx"); ap.add_argument("--slide", type=int, required=True); ap.add_argument("--box", type=int, required=True)
    ap.add_argument("--info", action="store_true"); ap.add_argument("--image"); ap.add_argument("--kind", default="diagram"); ap.add_argument("--out")
    a = ap.parse_args()
    prs = Presentation(a.pptx); sl = prs.slides[a.slide - 1]; bs = boxes_of(sl)
    if not bs or a.box > len(bs): sys.exit(f"{a.slide}장에 박스 {a.box}가 없음 (박스 수 {len(bs)})")
    box = bs[a.box - 1]; A, msg = inner_of(sl, box)
    targets = [sh for sh in sl.shapes if sh is not box and sh not in msg and inside(sh, A) and not (abs(sh.top / EMU - BY) < 0.03 and abs(sh.height / EMU - TITLE_H) < 0.03)]
    content = "\n".join(t for t in (describe(sh) for sh in targets) if t)
    if a.info or not a.image:
        print(f"{a.slide}장 박스 {a.box} 안쪽 영역: x={A[0]:.2f} y={A[1]:.2f} w={A[2]:.2f} h={A[3]:.2f} (인치)")
        print(f"→ HTML은 render_html.py --fit {A[2]:.2f} {A[3]:.2f} 로 그린다")
        print(f"박스 메시지: {msg[0].text_frame.text if msg else '(없음)'}\n현재 내용(도형 {len(targets)}개):\n{content}")
        return
    for sh in targets: sh._element.getparent().remove(sh._element)
    pic = sl.shapes.add_picture(a.image, Emu(int(A[0] * EMU)), Emu(int(A[1] * EMU)), Emu(int(A[2] * EMU)), Emu(int(A[3] * EMU)))
    pic.name = f"crp:{a.kind}"
    nt = sl.notes_slide.notes_text_frame
    nt.text = (nt.text + "\n" if nt.text else "") + f"[업그레이드 전 원본 — 박스 {a.box}]\n{content}"
    out = a.out or a.pptx; prs.save(out)
    print(f"교체 완료: {a.slide}장 박스 {a.box} → {a.image} ({A[2]:.2f}×{A[3]:.2f}in), 원본 내용은 노트에 보존 → {out}")

if __name__ == "__main__":
    main()
