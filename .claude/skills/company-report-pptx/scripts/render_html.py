"""HTML 시각화(도식·차트·설명 그림)를 고해상도 PNG로 렌더링 → PPT 박스에 삽입용
사용법: python render_html.py 입력.html 출력.png [--width 1200] [--height 700] [--scale 3]
- 브라우저(Playwright Chromium)로 HTML을 열어 그대로 캡처한다. CSS·SVG·JS 시각화 모두 가능.
- 외부 CDN 스크립트는 사내망에서 막힐 수 있으니 JS 라이브러리는 로컬 파일로 넣거나 순수 SVG/CSS로 그린다.
- 결과 PNG의 원본 데이터는 슬라이드 노트에 남긴다(이미지는 수정 불가).
- 숫자가 핵심인 차트는 PPT 기본 차트를 우선한다. 이 스크립트는 개념도·흐름도·원리 설명 그림용.
"""
import sys as _sys
for _s in (_sys.stdout, _sys.stderr):
    try: _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception: pass

import sys, argparse
from pathlib import Path

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("html"); ap.add_argument("png")
    ap.add_argument("--width", type=int, default=1200); ap.add_argument("--height", type=int, default=700)
    ap.add_argument("--scale", type=float, default=3)
    ap.add_argument("--fit", nargs=2, type=float, metavar=("W_IN", "H_IN"), help="PPT에 넣을 영역 크기(인치). 이 크기 그대로 그려서 1:1로 넣으면 글자 크기가 PPT 본문과 일치(1pt = 1.333px)")
    ap.add_argument("--channel", default=None, help="사내 PC에서 브라우저 다운로드가 막히면 msedge 또는 chrome (설치된 브라우저 사용)")
    a = ap.parse_args()
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        b = p.chromium.launch(channel=a.channel) if a.channel else p.chromium.launch()
        if a.fit: a.width, a.height = round(a.fit[0] * 96), round(a.fit[1] * 96)
        pg = b.new_page(viewport={"width": a.width, "height": a.height}, device_scale_factor=a.scale)
        pg.goto(Path(a.html).resolve().as_uri()); pg.wait_for_timeout(500)
        if a.fit:  # 영역 크기에 정확히 맞춤: #figure를 뷰포트 전체로
            pg.add_style_tag(content="html,body{margin:0;width:100%;height:100%} #figure{width:100%;height:100%;box-sizing:border-box}")
            pg.wait_for_timeout(100)
        el = pg.query_selector("#figure") or pg.query_selector("body")
        el.screenshot(path=a.png, omit_background=True)
        b.close()
    print("saved", a.png)

if __name__ == "__main__":
    main()
