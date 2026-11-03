"""결과 PPT를 장별 PNG로 렌더링 (검수용). Claude Code·Gemini CLI 공용.
사용법: python render_pptx.py 결과.pptx [출력폴더] [--dpi 80]
- Windows + PowerPoint 설치: PowerPoint로 직접 내보냄(실제 화면과 가장 같음, pywin32 필요)
- 그 외: LibreOffice(soffice)로 PDF 변환 후 PNG
출력: 출력폴더/slide-01.png ... 와 전체를 한 장에 모은 contact.png (여러 장을 한 번에 보기용)
"""
import sys as _sys
for _s in (_sys.stdout, _sys.stderr):
    try: _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception: pass

import sys, os, subprocess, shutil, glob, argparse, tempfile
from pathlib import Path

def via_powerpoint(pptx, out, width):
    import win32com.client  # pywin32
    app = win32com.client.Dispatch("PowerPoint.Application")
    pres = app.Presentations.Open(str(Path(pptx).resolve()), WithWindow=False)
    for i, s in enumerate(pres.Slides, 1):
        s.Export(str(Path(out, f"slide-{i:02d}.png").resolve()), "PNG", width, int(width * 9 / 16))
    pres.Close(); app.Quit()

def via_soffice(pptx, out, dpi):
    exe = shutil.which("soffice") or shutil.which("libreoffice") or next((p for p in [r"C:\Program Files\LibreOffice\program\soffice.exe", r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"] if os.path.exists(p)), None)
    if not exe: raise RuntimeError("PowerPoint(pywin32)도 LibreOffice도 없음 — 둘 중 하나 설치 필요")
    tmp = tempfile.mkdtemp()
    subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir", tmp, pptx], check=True, capture_output=True)
    pdf = glob.glob(os.path.join(tmp, "*.pdf"))[0]
    try:
        try: import pymupdf as fitz
        except ImportError: import fitz  # 구버전 PyMuPDF
        doc = fitz.open(pdf)
        for i, p in enumerate(doc, 1): p.get_pixmap(dpi=dpi).save(os.path.join(out, f"slide-{i:02d}.png"))
    except ImportError:
        subprocess.run(["pdftoppm", "-png", "-r", str(dpi), pdf, os.path.join(out, "slide")], check=True)
        for f in glob.glob(os.path.join(out, "slide-*.png")):  # pdftoppm 번호 자리수 통일
            n = int(Path(f).stem.split("-")[1]); os.replace(f, os.path.join(out, f"slide-{n:02d}.png"))

REVIEW_W = 1280   # 검수용 이미지 가로(px). 큰 PNG를 여러 장 한꺼번에 넣으면 일부 AI 도구(Gemini CLI 등)에서 400 오류(입력 초과)
SHEET_W = 1600    # 모아보기 한 장 가로(px). 2×2 = 4장씩

def review_images(out):
    """장별 작은 JPG(review/slide-XX.jpg) 생성 — 검수는 이걸 본다"""
    from PIL import Image
    rd = os.path.join(out, "review"); os.makedirs(rd, exist_ok=True)
    for f in sorted(glob.glob(os.path.join(out, "slide-*.png"))):
        im = Image.open(f).convert("RGB")
        if im.width > REVIEW_W: im = im.resize((REVIEW_W, int(im.height * REVIEW_W / im.width)))
        im.save(os.path.join(rd, Path(f).stem + ".jpg"), quality=82)

def sheets(out, nums, prefix="sheet"):
    """지정한 장들을 4장씩 2×2로 모은 JPG 여러 개 → [(파일, [장 번호])]"""
    from PIL import Image, ImageDraw, ImageFont
    res = []
    try: font = ImageFont.truetype("malgun.ttf", 22)
    except Exception:
        try: font = ImageFont.truetype("/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc", 22)
        except Exception: font = ImageFont.load_default()
    for k in range(0, len(nums), 4):
        grp = nums[k:k + 4]; cw = SHEET_W // 2
        ims = [Image.open(os.path.join(out, f"slide-{i:02d}.png")).convert("RGB") for i in grp]
        ch = int(ims[0].height * cw / ims[0].width); band = 40   # 장 번호는 이미지 위 띠에(슬라이드 내용을 가리지 않게)
        rows = (len(grp) + 1) // 2
        sh = Image.new("RGB", (SHEET_W, (ch + band) * rows), "white"); d = ImageDraw.Draw(sh)
        for j, (i, im) in enumerate(zip(grp, ims)):
            x, y = (j % 2) * cw, (j // 2) * (ch + band)
            d.rectangle((x + 3, y + 2, x + cw - 3, y + band - 4), fill=(31, 56, 100)); d.text((x + 12, y + 6), f"slide-{i:02d}", fill="white", font=font)
            sh.paste(im.resize((cw - 6, ch - 6)), (x + 3, y + band))
            d.rectangle((x + 3, y + band, x + cw - 4, y + band + ch - 7), outline=(160, 160, 160), width=1)
        name = os.path.join(out, f"{prefix}-{k // 4 + 1}.jpg"); sh.save(name, quality=80); res.append((name, grp))
    return res

def contact(out):
    from PIL import Image
    fs = sorted(glob.glob(os.path.join(out, "slide-*.png")))
    ims = [Image.open(f).convert("RGB") for f in fs]
    if not ims: return
    w, h = ims[0].size; cols = 2; rows = (len(ims) + 1) // 2
    sheet = Image.new("RGB", (w * cols, h * rows), "white")
    for k, im in enumerate(ims): sheet.paste(im.resize((w, h)), ((k % cols) * w, (k // cols) * h))
    sheet.save(os.path.join(out, "contact.png"))

def main():
    ap = argparse.ArgumentParser(); ap.add_argument("pptx"); ap.add_argument("out", nargs="?", default="render")
    ap.add_argument("--dpi", type=int, default=80); a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    for f in glob.glob(os.path.join(a.out, "slide-*.png")) + glob.glob(os.path.join(a.out, "*sheet-*.jpg")): os.remove(f)
    done = False
    if os.name == "nt":
        try: via_powerpoint(a.pptx, a.out, int(13.33 * a.dpi)); done = True
        except Exception as e: print("PowerPoint 내보내기 실패, LibreOffice로 시도:", e)
    if not done: via_soffice(a.pptx, a.out, a.dpi)
    contact(a.out); review_images(a.out)
    n = len(glob.glob(os.path.join(a.out, "slide-*.png")))
    sh = sheets(a.out, list(range(1, n + 1)))
    print("렌더링 완료:", a.out, f"({n}장) — 검수용: review/slide-XX.jpg, 4장 모아보기: " + ", ".join(Path(f).name for f, _ in sh))

if __name__ == "__main__":
    main()
