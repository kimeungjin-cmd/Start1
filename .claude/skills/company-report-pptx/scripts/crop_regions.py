"""손필기·스캔 이미지에서 구역(그래프·그림·수식·회로 스케치)을 잘라내고, 확인용 겹침 이미지를 만든다.
사용법: python crop_regions.py 원본.png(또는 .jpg, PDF는 먼저 쪽 이미지로 렌더링) 구역.json 출력폴더 [--enhance]
구역.json 예 (좌표는 이미지 대비 %, 왼쪽 위 기준):
[{"id":"G1","type":"그래프","x":5,"y":40,"w":45,"h":30,"note":"V-t 파형"},
 {"id":"C1","type":"회로","x":55,"y":10,"w":40,"h":35}]
출력: 출력폴더/G1.png ... 와 overlay.png(원본 위에 구역 상자·번호 표시)
→ overlay.png를 반드시 열어 상자가 대상을 빠짐없이 감싸는지 확인하고, 잘리면 좌표를 고쳐 다시 실행한다.
--enhance: 연필 필기처럼 흐린 원본의 대비를 올려서 잘라냄(원본 파일은 그대로 둠)
"""
import sys as _sys
for _s in (_sys.stdout, _sys.stderr):
    try: _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception: pass

import json, sys, argparse
from pathlib import Path
from PIL import Image, ImageDraw, ImageOps, ImageEnhance, ImageFont

def main():
    ap = argparse.ArgumentParser(); ap.add_argument("img"); ap.add_argument("regions"); ap.add_argument("out"); ap.add_argument("--enhance", action="store_true")
    a = ap.parse_args(); out = Path(a.out); out.mkdir(parents=True, exist_ok=True)
    im = ImageOps.exif_transpose(Image.open(a.img)).convert("RGB")  # 휴대폰 사진 회전 정보 반영
    W, H = im.size; src = im
    if a.enhance: src = ImageEnhance.Contrast(ImageOps.autocontrast(im.convert("L"), cutoff=2)).enhance(1.6).convert("RGB")
    ov = im.copy(); d = ImageDraw.Draw(ov)
    try: font = ImageFont.truetype("malgun.ttf", max(18, W // 60))
    except Exception:
        try: font = ImageFont.truetype("/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc", max(18, W // 60))
        except Exception: font = ImageFont.load_default()
    for r in json.loads(Path(a.regions).read_text(encoding="utf-8")):
        x0, y0 = int(W * r["x"] / 100), int(H * r["y"] / 100); x1, y1 = int(W * (r["x"] + r["w"]) / 100), int(H * (r["y"] + r["h"]) / 100)
        x0, y0, x1, y1 = max(0, x0), max(0, y0), min(W, x1), min(H, y1)
        src.crop((x0, y0, x1, y1)).save(out / f"{r['id']}.png")
        d.rectangle((x0, y0, x1, y1), outline=(192, 0, 0), width=max(3, W // 300))
        d.text((x0 + 6, y0 + 4), f"{r['id']} {r.get('type','')}", fill=(192, 0, 0), font=font)
        print(f"{r['id']}\t{r.get('type','')}\t{x1-x0}x{y1-y0}\t{out / (r['id'] + '.png')}")
    ov.save(out / "overlay.png"); print("확인용:", out / "overlay.png")

if __name__ == "__main__": main()
