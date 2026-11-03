"""원본 문서의 사진·그림을 번호(P1, P2…)를 붙여 꺼내고, 목록(assets_index.md)과 번호 붙은 모아보기 이미지를 만든다.
웹의 상위 모델에 원본을 분석시킬 때 md 스토리보드에는 그림을 담을 수 없으므로, 그림은 '번호로' 주고받는다.

사용법: python extract_assets.py 원본파일 [출력폴더=assets] [--min 120]
지원: .pdf .pptx .docx .xlsx .png/.jpg(그대로 1장)
  PDF: 사진(이미지) + 선·도형으로 그린 그림(데이터시트 그래프·회로도·블록도)도 잡음. 표 테두리만 있는 묶음은 제외. 큰 문서는 --pages 1-5,250-252
  PPT: 그룹 안 그림, PPT에서 잘라 쓴(crop) 상태·회전을 화면 그대로 반영. 위치는 '3장#2'(그 장의 2번째 그림)
출력:
  assets/P01.png …           꺼낸 그림 (엑셀·PPT의 회전 설정 반영, PDF 조각 그림은 쪽 전체로 대체)
  assets/page-01.png …       PDF는 쪽 전체 이미지도 같이 (조각 이미지·벡터 그래프 대비)
  assets/assets_index.md     번호 / 위치(쪽·장·행) / 크기 / 근처 글  ← 웹 모델에 원본과 같이 올린다
  assets/assets-sheet-N.jpg  번호가 적힌 모아보기(12개씩) ← 웹 모델에 같이 올리면 번호로 지목 가능
웹 모델은 스토리보드에 '사진=P03' 처럼 번호로만 적고, 실행 쪽(이 스킬)이 assets/의 실제 파일을 넣는다.
"""
import sys as _sys
for _s in (_sys.stdout, _sys.stderr):
    try: _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception: pass
import argparse, io, os, re, zipfile
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont, ImageOps

def save(im, out, n, rows, where, near=""):
    pid = f"P{n:02d}"; im = im.convert("RGB"); im.save(out / f"{pid}.png")
    rows.append((pid, where, f"{im.width}×{im.height}", re.sub(r"\s+", " ", near).strip()[:80])); return n + 1

def from_pdf(src, out, rows, mn):
    """쪽마다 판단: ①글이 거의 없는 쪽(스캔·캡처 쪽) → 쪽 전체를 그림 하나로 ②큰 그림 → 그 영역만 잘라냄
    ③글 쪽의 자잘한 그림 조각(강조 표시 등) → 그림으로 세지 않고 목록에 메모만"""
    try: import pymupdf as fitz
    except ImportError: import fitz
    doc = fitz.open(src); n = 1
    pages = PAGES or set(range(1, len(doc) + 1))
    for pi, fp in enumerate(doc, 1):
        if pi not in pages: continue
        pix = fp.get_pixmap(dpi=110); pim = Image.open(io.BytesIO(pix.tobytes("png"))); pim.save(out / f"page-{pi:02d}.png")
        text = fp.get_text().strip(); info = fp.get_image_info(); W, H = fp.rect.width, fp.rect.height; s = 110 / 72
        if not info:
            for y0, x0, crop, where, near in sorted(vector_figs(fp, pim, pi, s, W, H, []), key=lambda c: (round(c[0] / 20), c[1])):
                n = save(crop, out, n, rows, where, near)
            continue
        if len(text) < 120:
            n = save(pim, out, n, rows, f"{pi}쪽 전체(그림·캡처 쪽)", text[:80]); continue
        big = [x for x in info if (x["bbox"][2] - x["bbox"][0]) * (x["bbox"][3] - x["bbox"][1]) >= 0.05 * W * H]
        cands = []
        for x in big:
            x0, y0, x1, y1 = x["bbox"]; crop = pim.crop((int(x0 * s), int(y0 * s), int(x1 * s), int(y1 * s)))
            if min(crop.size) < mn / 2: continue
            near = fp.get_textbox(fitz.Rect(0, max(0, y0 - 60), W, y0)) + " / " + fp.get_textbox(fitz.Rect(0, y1, W, min(H, y1 + 60)))
            cands.append((y0, x0, crop, f"{pi}쪽 사진 ({int(x0)},{int(y0)}pt)", near))
        cands += vector_figs(fp, pim, pi, s, W, H, big)
        for y0, x0, crop, where, near in sorted(cands, key=lambda c: (round(c[0] / 20), c[1])):  # 위→아래, 왼→오른쪽
            n = save(crop, out, n, rows, where, near)
        small = len(info) - len(big)
        if small >= 3 and not big:
            rows.append(("—", f"{pi}쪽", f"작은 조각 {small}개", "강조 표시 등 장식으로 판단 — 필요하면 page-%02d.png 참고" % pi))
    return n
def vector_figs(fp, pim, pi, s, W, H, photos):
    """선·도형으로 그린 그림(데이터시트 그래프·회로도·블록도 등)을 묶어서 잘라냄. 표 테두리처럼 가는 선만 있는 묶음은 제외"""
    try: dr = fp.get_drawings()
    except Exception: return []
    items = [(d["rect"], (d["rect"].width < 2.5 or d["rect"].height < 2.5) and not d.get("fill"))
             for d in dr if d.get("rect") is not None and d["rect"].width * d["rect"].height < 0.6 * W * H]
    rects = [r for r, _ in items]
    if len(rects) < 4: return []
    import pymupdf as _f
    groups = []
    for r in rects:
        r = _f.Rect(r.x0 - 8, r.y0 - 8, r.x1 + 8, r.y1 + 8)
        for g in groups:
            if g.intersects(r): g |= r; break
        else: groups.append(r)
    merged = True
    while merged:
        merged = False
        for i in range(len(groups)):
            for j in range(i + 1, len(groups)):
                if groups[i].intersects(groups[j]): groups[i] |= groups[j]; groups.pop(j); merged = True; break
            if merged: break
    res = []
    for g in groups:
        g &= fp.rect
        if any(_f.Rect(p["bbox"]).intersects(g) for p in photos): continue  # 사진 테두리 등은 사진으로 이미 잡음
        if g.width * g.height < 0.03 * W * H or g.width < 80 or g.height < 60: continue
        if g.width * g.height > 0.85 * W * H: continue  # 쪽 전체 테두리 등
        mem = [thin for r, thin in items if g.contains(r)]
        if len(mem) < 3 or all(mem): continue   # 가는 선만 있는 묶음 = 표 테두리·밑줄 → 제외
        crop = pim.crop((int(g.x0 * s), int(g.y0 * s), int(g.x1 * s), int(g.y1 * s)))
        near = fp.get_textbox(_f.Rect(0, max(0, g.y0 - 40), W, g.y0)).strip()
        res.append((g.y0, g.x0, crop, f"{pi}쪽 그림(선·도형) ({int(g.x0)},{int(g.y0)}pt)", near))
    return res

def from_pptx(src, out, rows, mn):
    """그룹 안의 그림까지 찾고, PPT에서 잘라 쓴(crop) 상태와 회전을 화면 그대로 반영"""
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    prs = Presentation(src); n = 1
    def walk(shapes):
        for sh in shapes:
            if sh.shape_type == MSO_SHAPE_TYPE.GROUP: yield from walk(sh.shapes)
            elif sh.shape_type == MSO_SHAPE_TYPE.PICTURE or (getattr(sh, "image", None) is not None and sh.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER):
                yield sh
    for si, sl in enumerate(prs.slides, 1):
        texts = " ".join(sh.text_frame.text for sh in sl.shapes if sh.has_text_frame)[:80]
        for k, sh in enumerate(sorted(walk(sl.shapes), key=lambda x: (round((x.top or 0) / 91440), x.left or 0)), 1):  # 위→아래, 왼→오른쪽
            try: im = Image.open(io.BytesIO(sh.image.blob))
            except Exception: continue
            if min(im.size) < mn: continue
            try:
                l, t, r, b = sh.crop_left, sh.crop_top, sh.crop_right, sh.crop_bottom
                if any(v for v in (l, t, r, b)):
                    W, H = im.size; im = im.crop((int(W * l), int(H * t), int(W * (1 - r)), int(H * (1 - b))))
            except Exception: pass
            if sh.rotation: im = im.rotate(-sh.rotation, expand=True)
            n = save(im, out, n, rows, f"{si}장#{k}", texts)
    return n

def from_docx(src, out, rows, mn):
    import docx
    d = docx.Document(src); n = 1; last = ""
    rels = {r.rId: r for r in d.part.rels.values() if "image" in r.reltype}
    for pi, p in enumerate(d.paragraphs, 1):
        if p.text.strip(): last = p.text
        for blip in p._p.xpath(".//a:blip"):
            rid = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
            if rid in rels:
                im = Image.open(io.BytesIO(rels[rid].target_part.blob))
                if min(im.size) >= mn: n = save(im, out, n, rows, f"{pi}번째 문단 근처", last)
    return n

def from_xlsx(src, out, rows, mn):
    import openpyxl
    wb = openpyxl.load_workbook(src); n = 1
    z = zipfile.ZipFile(src); rot = {}
    for name in z.namelist():   # 도형 회전 설정(엑셀 화면 방향) 반영
        if name.startswith("xl/drawings/drawing") and name.endswith(".xml"):
            x = z.read(name).decode("utf-8", "ignore")
            for k, m in enumerate(re.finditer(r"<xdr:(?:twoCellAnchor|oneCellAnchor).*?</xdr:(?:twoCellAnchor|oneCellAnchor)>", x, re.S)):
                r = re.search(r'rot="(-?\d+)"', m.group(0)); rot[(name, k)] = int(r.group(1)) / 60000 if r else 0
    for ws in wb:
        dname = None
        for k, im0 in enumerate(ws._images):
            im = Image.open(io.BytesIO(im0._data()))
            if min(im.size) < mn: continue
            r0 = im0.anchor._from.row + 1
            ang = next((v for (nm, kk), v in rot.items() if kk == k), 0)
            if ang: im = im.rotate(-ang, expand=True)
            near = " ".join(str(c.value) for c in ws[r0] if c.value is not None)[:80]
            n = save(im, out, n, rows, f"{ws.title} {r0}행", near)
    return n

def sheet(out, rows):
    try: font = ImageFont.truetype("malgun.ttf", 22)
    except Exception:
        try: font = ImageFont.truetype("/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc", 22)
        except Exception: font = ImageFont.load_default()
    T = 300; files = []
    for g in range(0, len(rows), 12):
        grp = rows[g:g + 12]; cols = 4; rr = (len(grp) + cols - 1) // cols
        sh = Image.new("RGB", (cols * (T + 10), rr * (T + 46)), "white"); d = ImageDraw.Draw(sh)
        for k, (pid, *_ ) in enumerate([r for r in grp if r[0] != "—"]):
            im = Image.open(out / f"{pid}.png"); im.thumbnail((T, T)); x, y = (k % cols) * (T + 10), (k // cols) * (T + 46)
            d.rectangle((x, y, x + 90, y + 36), fill=(192, 0, 0)); d.text((x + 10, y + 5), pid, fill="white", font=font)
            sh.paste(im, (x, y + 42))
        f = out / f"assets-sheet-{g // 12 + 1}.jpg"; sh.save(f, quality=80); files.append(f.name)
    return files

def find(out, query):
    """스토리보드의 '원본 위치'(예: 4쪽, 3장, 12행, 5쪽#2)를 그림 번호로 바꿔 준다"""
    rows = [l.split("|")[1:4] for l in (Path(out) / "assets_index.md").read_text(encoding="utf-8").splitlines() if l.startswith("| P")]
    q, k = (query.split("#") + ["0"])[:2]; q = q.strip(); num = re.sub(r"\D", "", q); unit = re.sub(r"[\d\s]", "", q)
    hits = [r[0].strip() for r in rows if re.search(rf"(^|\D){num}\s*{unit}", r[1])]
    if k != "0" and hits: hits = hits[int(k) - 1:int(k)] if int(k) <= len(hits) else []
    print(f"{query} → {', '.join(hits) if hits else '해당 그림 없음 (assets_index.md 확인)'}")
    return hits

def main():
    if len(sys.argv) >= 3 and sys.argv[1] == "--find":   # python extract_assets.py --find "4쪽" [assets]
        find(sys.argv[3] if len(sys.argv) > 3 else "assets", sys.argv[2]); return
    ap = argparse.ArgumentParser(); ap.add_argument("src"); ap.add_argument("out", nargs="?", default="assets"); ap.add_argument("--min", type=int, default=120)
    ap.add_argument("--pages", help="PDF 쪽 범위 예: 1-5,12,250-252 (큰 문서용)")
    a = ap.parse_args(); out = Path(a.out); out.mkdir(parents=True, exist_ok=True); rows = []
    global PAGES
    if a.pages:
        PAGES = set()
        for part in a.pages.split(","):
            x, _, y = part.partition("-"); PAGES |= set(range(int(x), int(y or x) + 1))
    ext = Path(a.src).suffix.lower()
    if ext == ".pdf": from_pdf(a.src, out, rows, a.min)
    elif ext == ".pptx": from_pptx(a.src, out, rows, a.min)
    elif ext == ".docx": from_docx(a.src, out, rows, a.min)
    elif ext in (".xlsx", ".xlsm"): from_xlsx(a.src, out, rows, a.min)
    elif ext in (".png", ".jpg", ".jpeg"): save(ImageOps.exif_transpose(Image.open(a.src)), out, 1, rows, "원본 이미지 1장")
    else: sys.exit("지원 형식: pdf pptx docx xlsx png jpg")
    sheets = sheet(out, [r for r in rows if r[0] != "—"]) if any(r[0] != "—" for r in rows) else []
    md = [f"# 원본 그림 목록: {Path(a.src).name}", "", "| 번호 | 위치 | 크기(px) | 근처 글 |", "|---|---|---|---|"]
    md += [f"| {p} | {w} | {s} | {t} |" for p, w, s, t in rows]
    md += ["", f"모아보기: {', '.join(sheets) if sheets else '없음'}",
           "스토리보드에는 그림을 '사진=P03'처럼 번호로 적는다. 원본에 없는 그림 번호를 만들지 않는다.",
           "그래프 이미지의 숫자를 읽어 새 차트로 다시 그리지 않는다(필요하면 원본 그림 그대로 사용)."]
    (out / "assets_index.md").write_text("\n".join(md), encoding="utf-8")
    print(f"그림 {sum(1 for r in rows if r[0] != '—')}개 → {out}/ (목록 assets_index.md, 모아보기 {', '.join(sheets)})")

PAGES = None
import sys
if __name__ == "__main__":
    main()
