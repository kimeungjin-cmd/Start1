"""원본 대비 결과 PPT의 수치·사진 누락 검사.
사용법: python check_coverage.py <결과.pptx> <원본1> [원본2 ...]
원본 형식: .pptx .docx .pdf .xlsx .txt .md .eml(텍스트)
- 수치: 원본에 있는 숫자가 결과 PPT(본문·주석·유첨·노트 포함)에 없으면 '누락 후보'로 출력
- 사진: 원본 이미지 개수 vs 결과 이미지 개수 비교 (작은 아이콘·로고는 제외)
누락 후보가 있으면 종료코드 1. 의도적으로 제외한 숫자(쪽 번호·문서 번호 등)는 --ignore 로 지정.
"""
import sys as _sys
for _s in (_sys.stdout, _sys.stderr):
    try: _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception: pass

import re, sys, argparse
from pathlib import Path

NUM = re.compile(r"(?<![0-9A-Za-z.,])[-+]?\d{1,3}(?:,\d{3})+(?:\.\d+)?|(?<![0-9A-Za-z.,])[-+]?\d+(?:\.\d+)?")  # 한글 바로 뒤 숫자(제81조)도 잡음
MIN_IMG_PX = 120  # 이보다 작은 이미지는 아이콘·로고로 보고 제외

def norm(n):
    n = n.replace(",", "").lstrip("+")
    if "." in n:
        n = n.rstrip("0").rstrip(".")
    return n

def pptx_text_imgs(path):
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    prs = Presentation(path); texts, imgs = [], 0
    def walk(shapes):
        nonlocal imgs
        for sh in shapes:
            if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
                walk(sh.shapes); continue
            if sh.has_text_frame: texts.append(sh.text_frame.text)
            if getattr(sh, "has_table", False) and sh.has_table:
                for r in sh.table.rows:
                    for c in r.cells: texts.append(c.text)
            if getattr(sh, "has_chart", False) and sh.has_chart:
                for p in sh.chart.plots:
                    for se in p.series: texts.extend(str(v) for v in se.values if v is not None)
                    texts.extend(str(c) for c in p.categories)
            if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    from PIL import Image; import io
                    w, h = Image.open(io.BytesIO(sh.image.blob)).size
                    if min(w, h) >= MIN_IMG_PX: imgs += 1
                except Exception: imgs += 1
    for sl in prs.slides:
        walk(sl.shapes)
        if sl.has_notes_slide: texts.append(sl.notes_slide.notes_text_frame.text)
    return "\n".join(texts), imgs

def docx_text_imgs(path):
    import docx, zipfile
    d = docx.Document(path)
    t = [p.text for p in d.paragraphs]
    for tb in d.tables:
        for r in tb.rows:
            for c in r.cells: t.append(c.text)
    imgs = len([n for n in zipfile.ZipFile(path).namelist() if n.startswith("word/media/")])
    return "\n".join(t), imgs

def pdf_text_imgs(path):
    import pdfplumber
    t, imgs = [], 0
    with pdfplumber.open(path) as pdf:
        for pg in pdf.pages:
            t.append(pg.extract_text() or "")
            imgs += len([im for im in pg.images if min(im["width"], im["height"]) >= 40])
    return "\n".join(t), imgs

def xlsx_text_imgs(path):
    import openpyxl, zipfile
    wb = openpyxl.load_workbook(path, data_only=True); t = []
    for ws in wb:
        for row in ws.iter_rows(values_only=True):
            t += [str(v) for v in row if v is not None and not hasattr(v, "year")]  # 날짜 셀은 제외(형식 차이)
    imgs = len([n for n in zipfile.ZipFile(path).namelist() if n.startswith("xl/media/")])
    return "\n".join(t), imgs

def load(path):
    p = Path(path); ext = p.suffix.lower()
    if ext == ".pptx": return pptx_text_imgs(p)
    if ext == ".docx": return docx_text_imgs(p)
    if ext == ".pdf": return pdf_text_imgs(p)
    if ext in (".xlsx", ".xlsm"): return xlsx_text_imgs(p)
    return p.read_text(encoding="utf-8", errors="ignore"), 0

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("output"); ap.add_argument("sources", nargs="+")
    ap.add_argument("--ignore", nargs="*", default=[], help="검사에서 뺄 숫자")
    a = ap.parse_args()
    out_text, out_imgs = load(a.output)
    out_nums = {norm(n) for n in NUM.findall(out_text)}
    src_nums, src_imgs = {}, 0
    for sp in a.sources:
        t, im = load(sp); src_imgs += im
        for n in NUM.findall(t): src_nums.setdefault(norm(n), sp)
    ignore = {norm(x) for x in a.ignore}
    missing = sorted([n for n in src_nums if n not in out_nums and n not in ignore], key=lambda x: (len(x), x))
    print(f"[수치] 원본 고유 숫자 {len(src_nums)}개 / 결과에 없는 숫자 {len(missing)}개")
    for n in missing: print(f"  - 누락 후보: {n}  (원본: {Path(src_nums[n]).name})")
    print(f"[사진] 원본 {src_imgs}장 / 결과 {out_imgs}장" + ("  ← 결과가 적음, 누락 확인 필요" if out_imgs < src_imgs else ""))
    sys.exit(1 if missing or out_imgs < src_imgs else 0)

if __name__ == "__main__":
    main()
