"""수식(LaTeX)을 고해상도 PNG로 렌더링 → PPT에 삽입용 (KaTeX + Playwright, 인터넷 불필요)
사용법:
  python render_math.py --latex "f_c=\\frac{1}{2\\pi RC}" --out eq1.png
  python render_math.py --file 수식.json --outdir eq/      # [{"id":"E1","latex":"..."}, ...]
옵션: --size 28(px) --color 222222 --scale 3
- 결과 PNG와 함께 크기(px)를 출력한다 → D.photo(..., [w,h])에 그대로 사용
- 원본 LaTeX는 슬라이드 노트에 반드시 남긴다(이미지 수식은 직접 수정 불가)
"""
import sys as _sys
for _s in (_sys.stdout, _sys.stderr):
    try: _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception: pass

import json, argparse, html, sys
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
KATEX = ROOT / "node_modules" / "katex" / "dist"

def page(latex, size, color):
    return f"""<!DOCTYPE html><html><head><meta charset='utf-8'>
<link rel='stylesheet' href='{(KATEX / "katex.min.css").as_uri()}'>
<script src='{(KATEX / "katex.min.js").as_uri()}'></script>
<style>body{{margin:0;background:transparent}}#figure{{display:inline-block;padding:6px 10px;font-size:{size}px;color:#{color}}}</style>
</head><body><div id='figure'></div><script>
try {{ katex.render({json.dumps(latex)}, document.getElementById('figure'), {{displayMode:true, throwOnError:true}}); }}
catch(e) {{ const d=document.createElement('span'); d.className='katex-error'; d.title=e.message; document.body.appendChild(d); }}
</script></body></html>"""

def render(items, size, color, scale):
    if not (KATEX / "katex.min.js").exists():
        sys.exit(f"KaTeX 없음 → 스킬 폴더에서 npm install ({ROOT})")
    from playwright.sync_api import sync_playwright
    from PIL import Image
    out = []
    with sync_playwright() as p:
        try: b = p.chromium.launch()
        except Exception: b = p.chromium.launch(channel="msedge")
        pg = b.new_page(viewport={"width": 1600, "height": 400}, device_scale_factor=scale)
        for it in items:
            tmp = Path(it["png"]).resolve().with_suffix(".html"); tmp.write_text(page(it["latex"], size, color), encoding="utf-8")
            pg.goto(tmp.as_uri()); pg.wait_for_timeout(200)
            err = pg.evaluate("document.querySelector('.katex-error') ? document.querySelector('.katex-error').title : ''")
            if err: print(f"[수식 오류] {it.get('id','')}: {err}"); tmp.unlink(); continue
            pg.query_selector("#figure .katex-display, #figure").screenshot(path=it["png"], omit_background=True)
            tmp.unlink()
            im = Image.open(it["png"]); bb = im.getchannel("A").getbbox()  # 투명 여백 잘라내기
            if bb: im.crop((max(bb[0]-8,0), max(bb[1]-8,0), min(bb[2]+8,im.width), min(bb[3]+8,im.height))).save(it["png"])
            w, h = Image.open(it["png"]).size; out.append((it.get("id", ""), it["png"], w, h))
            print(f"{it.get('id','')}\t{it['png']}\t{w}x{h}")
        b.close()
    return out

def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--latex"); ap.add_argument("--out", default="eq.png")
    ap.add_argument("--file"); ap.add_argument("--outdir", default="eq")
    ap.add_argument("--size", type=int, default=28); ap.add_argument("--color", default="222222"); ap.add_argument("--scale", type=float, default=3)
    a = ap.parse_args()
    if a.file:
        Path(a.outdir).mkdir(parents=True, exist_ok=True)
        items = [{"id": e["id"], "latex": e["latex"], "png": str(Path(a.outdir) / f"{e['id']}.png")} for e in json.loads(Path(a.file).read_text(encoding="utf-8"))]
    elif a.latex: items = [{"id": "", "latex": a.latex, "png": a.out}]
    else: sys.exit("--latex 또는 --file 필요")
    render(items, a.size, a.color, a.scale)

if __name__ == "__main__":
    main()
