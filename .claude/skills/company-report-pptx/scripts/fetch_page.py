"""자바스크립트로 그려지는 웹페이지를 실제 브라우저로 열어 본문을 가져온다 (검색·조사 보조)
일반 웹 가져오기(web fetch)는 화면을 그리기 전 HTML만 받아서, 요금표·운영시간·공지처럼
JS로 나중에 채워지는 내용이 비어 있는 경우가 많다. 그럴 때 이 스크립트를 쓴다.

사용법: python fetch_page.py URL [--wait 2500] [--selector "main"] [--screenshot shot.png] [--max 6000] [--click "더보기"]
- 출력: 제목, 최종 URL, 본문 텍스트(표는 탭 구분), 상위 링크 일부
- --selector: 본문 영역만 뽑기(없으면 body) / --click: 펼치기 버튼 글자를 눌러 숨은 내용 열기
- --screenshot: 화면 캡처(출처 증빙용·이미지로 확인할 때). 캡처 이미지를 보고서에 넣을 땐 출처 표기
- 공개 페이지만 대상. 로그인·결제·개인정보 페이지는 열지 않는다. 같은 사이트를 반복 요청하지 않는다.
언제 쓰나: web fetch 결과가 비었거나 짧거나 "JavaScript를 켜세요" 류 문구가 보일 때, 날짜·가격·운영시간이 안 보일 때
"""
import sys as _sys
for _s in (_sys.stdout, _sys.stderr):
    try: _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception: pass
import argparse

JS_TEXT = r"""(sel) => {
  const root = document.querySelector(sel) || document.body;
  root.querySelectorAll('script,style,noscript,svg').forEach(e => e.remove());
  root.querySelectorAll('table').forEach(t => {
    const rows = [...t.querySelectorAll('tr')].map(r => [...r.children].map(c => c.innerText.trim().replace(/\s+/g,' ')).join('\t'));
    const pre = document.createElement('pre'); pre.textContent = '\n[표]\n' + rows.join('\n') + '\n'; t.replaceWith(pre);
  });
  return root.innerText.replace(/\n{3,}/g, '\n\n').trim();
}"""

def main():
    ap = argparse.ArgumentParser(); ap.add_argument("url")
    ap.add_argument("--wait", type=int, default=2500); ap.add_argument("--selector", default="body")
    ap.add_argument("--screenshot"); ap.add_argument("--max", type=int, default=6000); ap.add_argument("--click")
    a = ap.parse_args()
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        try: b = p.chromium.launch()
        except Exception: b = p.chromium.launch(channel="msedge")
        pg = b.new_page(viewport={"width": 1366, "height": 900}, locale="ko-KR")
        pg.goto(a.url, wait_until="domcontentloaded", timeout=45000)
        try: pg.wait_for_load_state("networkidle", timeout=a.wait + 5000)
        except Exception: pass
        pg.wait_for_timeout(a.wait)
        if a.click:
            try: pg.get_by_text(a.click).first.click(timeout=3000); pg.wait_for_timeout(1500)
            except Exception: print(f"[알림] '{a.click}' 버튼을 찾지 못함")
        for _ in range(3): pg.mouse.wheel(0, 2500); pg.wait_for_timeout(400)  # 스크롤로 늦게 뜨는 내용 불러오기
        text = pg.evaluate(JS_TEXT, a.selector)
        links = pg.evaluate("() => [...document.querySelectorAll('a[href]')].slice(0,300).map(a=>[a.innerText.trim().slice(0,40), a.href]).filter(x=>x[0])")
        if a.screenshot: pg.screenshot(path=a.screenshot, full_page=True)
        print(f"# 제목: {pg.title()}\n# 최종 URL: {pg.url}\n# 본문 글자 수: {len(text)}\n")
        print(text[:a.max] + ("\n…(이하 생략, --max로 늘릴 수 있음)" if len(text) > a.max else ""))
        print("\n# 링크(일부)"); [print(f"- {t} → {h}") for t, h in links[:25]]
        b.close()

if __name__ == "__main__":
    main()
