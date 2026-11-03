"""본문 박스 경계 침범 검사 (렌더링 전에 좌표로 1차 확인)
사용법: python check_layout.py <결과.pptx>
- 본문 박스(높이 약 5.05in, 위치 y≈1.95in) 안에서 시작한 도형이 박스 아래·오른쪽 경계를 넘으면 표시
- 슬라이드 밖으로 나간 도형도 표시
침범이 있으면 종료코드 1. 글자가 상자 안에서 넘치는 것은 렌더링으로 확인해야 함.
"""
import sys as _sys
for _s in (_sys.stdout, _sys.stderr):
    try: _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception: pass

import sys
from pptx import Presentation
EMU = 914400
TOL = 0.02  # 인치 허용오차

def main():
    prs = Presentation(sys.argv[1]); SW, SH = prs.slide_width / EMU, prs.slide_height / EMU
    issues = 0
    for n, sl in enumerate(prs.slides, 1):
        def real_h(sh):  # 표는 틀 높이가 실제와 다를 수 있어 행 높이 합으로 계산
            if getattr(sh, "has_table", False) and sh.has_table:
                return max(sh.height, sum(r.height for r in sh.table.rows)) / EMU
            return sh.height / EMU
        shapes = [(sh, sh.left / EMU, sh.top / EMU, sh.width / EMU, real_h(sh)) for sh in sl.shapes if sh.width is not None]
        boxes = [(x, y, w, h) for sh, x, y, w, h in shapes if abs(y - 1.95) < 0.03 and abs(h - 5.05) < 0.03]
        for sh, x, y, w, h in shapes:
            name = (sh.text_frame.text[:20] if sh.has_text_frame else sh.shape_type) if True else ""
            if x < -TOL or y < -TOL or x + w > SW + TOL or y + h > SH + TOL:
                print(f"[{n}장] 슬라이드 밖으로 나감: {name!r}"); issues += 1; continue
            for bx, by, bw, bh in boxes:
                if (bx, by, bw, bh) == (x, y, w, h):
                    continue
                inside = bx - TOL <= x <= bx + bw and by - TOL <= y <= by + bh
                if inside and (y + h > by + bh + TOL or x + w > bx + bw + TOL):
                    over = max(y + h - (by + bh), x + w - (bx + bw))
                    print(f"[{n}장] 본문 박스 침범 {over:.2f}in: {name!r}"); issues += 1
        # 겹침 검사: 표·차트·사진·채워진 도형끼리 일부만 겹치면 표시 (한쪽이 다른 쪽을 완전히 품는 배경/카드 구조는 정상)
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        solid = []
        for sh, x, y, w, h in shapes:
            if (x, y, w, h) in boxes or w < 0.05 or h < 0.05: continue
            kind = None
            if getattr(sh, "has_table", False) and sh.has_table: kind = "표"
            elif getattr(sh, "has_chart", False) and sh.has_chart: kind = "차트"
            elif sh.shape_type == MSO_SHAPE_TYPE.PICTURE: kind = "그림"
            elif sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                try:
                    pr = sh._element.xpath(".//a:prstGeom/@prst"); pr = pr[0] if pr else ""
                    if pr not in ("chevron", "homePlate") and sh.fill.type == 1 and w > 0.6 and h > 0.4: kind = "카드"  # 단계 화살표는 일부러 겹침
                except Exception: pass
            if kind: solid.append((kind, x, y, w, h))
        for i in range(len(solid)):
            for j in range(i + 1, len(solid)):
                k1, x1, y1, w1, h1 = solid[i]; k2, x2, y2, w2, h2 = solid[j]
                ix = min(x1 + w1, x2 + w2) - max(x1, x2); iy = min(y1 + h1, y2 + h2) - max(y1, y2)
                if ix <= TOL or iy <= TOL: continue
                inside = lambda ax, ay, aw, ah, bx, by, bw, bh: ax >= bx - TOL and ay >= by - TOL and ax + aw <= bx + bw + TOL and ay + ah <= by + bh + TOL
                if inside(x1, y1, w1, h1, x2, y2, w2, h2) or inside(x2, y2, w2, h2, x1, y1, w1, h1): continue
                print(f"[{n}장] {k1}와 {k2}가 겹침 ({ix:.2f}×{iy:.2f}in)"); issues += 1
    print(f"레이아웃 이슈 {issues}건"); sys.exit(1 if issues else 0)

if __name__ == "__main__":
    main()
