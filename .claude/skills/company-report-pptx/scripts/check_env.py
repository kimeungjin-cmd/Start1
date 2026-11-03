"""스킬 실행 환경 점검. 처음 설치 후, 또는 스크립트가 실패할 때 실행.
사용법: python scripts/check_env.py   (스킬 폴더에서)
"""
import sys as _sys
for _s in (_sys.stdout, _sys.stderr):
    try: _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception: pass

import shutil, subprocess, sys, importlib, os
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
ok = True
def mark(name, good, fix=""):
    global ok; ok &= good
    print(("[OK]   " if good else "[없음] ") + name + ("" if good else f"  → {fix}"))
node = shutil.which("node")
mark("Node.js", bool(node), "Node.js LTS 설치")
if node:
    r = subprocess.run([node, "-e", "require('pptxgenjs')"], cwd=ROOT / "scripts", capture_output=True)
    mark("pptxgenjs (스킬 폴더 기준)", r.returncode == 0, f"cd \"{ROOT}\" && npm install")
    mark("KaTeX (수식 렌더링)", (ROOT / "node_modules" / "katex" / "dist" / "katex.min.js").exists(), f"cd \"{ROOT}\" && npm install")
for mod, pipname in [("pptx", "python-pptx"), ("PIL", "pillow"), ("pdfplumber", "pdfplumber"), ("docx", "python-docx"), ("openpyxl", "openpyxl"), ("playwright", "playwright"), ("pymupdf", "pymupdf")]:
    try: importlib.import_module(mod); mark(pipname, True)
    except ImportError: mark(pipname, False, f"pip install -r \"{ROOT / 'requirements.txt'}\"")
pp = False
if os.name == "nt":
    try: importlib.import_module("win32com.client"); pp = True
    except ImportError: pass
lo = bool(shutil.which("soffice") or shutil.which("libreoffice") or os.path.exists(r"C:\Program Files\LibreOffice\program\soffice.exe"))
mark("렌더러 (PowerPoint+pywin32 또는 LibreOffice)", pp or lo, "Windows면 pip install pywin32 (PowerPoint 필요), 아니면 LibreOffice 설치")
try:
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        try: p.chromium.launch().close(); mark("Playwright 브라우저", True)
        except Exception:
            try: p.chromium.launch(channel="msedge").close(); mark("Playwright 브라우저 (Edge 사용 — render_html.py에 --channel msedge)", True)
            except Exception: mark("Playwright 브라우저", False, "python -m playwright install chromium")
except Exception: pass
print("\n환경 준비 완료" if ok else "\n위 [없음] 항목을 설치한 뒤 다시 실행")
sys.exit(0 if ok else 1)
