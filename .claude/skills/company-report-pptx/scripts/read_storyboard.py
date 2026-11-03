"""웹 모델이 만든 스토리보드를 워드(.docx)·PDF·텍스트 어느 형식으로 받아도 storyboard.md로 바꾼다.
사용법: python read_storyboard.py 받은파일(.docx|.pdf|.txt|.md) [storyboard.md]
- 워드의 표는 행마다 '칸 | 칸 | 칸'으로 옮긴다(웹에서 표로 정리했어도 내용이 빠지지 않게)
- 이후 SKILL.md의 '넘겨받기' 절차대로 진행
"""
import sys as _sys
for _s in (_sys.stdout, _sys.stderr):
    try: _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception: pass
import sys
from pathlib import Path

def from_docx(p):
    import docx
    d = docx.Document(p); out = []
    body = d.element.body
    for el in body.iterchildren():
        tag = el.tag.split("}")[1]
        if tag == "p":
            t = "".join(x.text or "" for x in el.iter() if x.tag.endswith("}t")).strip()
            style = el.xpath("string(.//w:pStyle/@w:val)") or ""
            if t: out.append(("## " if "Heading" in style or "제목" in style else "") + t)
        elif tag == "tbl":
            for tr in el.iter():
                if tr.tag.endswith("}tr"):
                    cells = ["".join(x.text or "" for x in tc.iter() if x.tag.endswith("}t")).strip() for tc in tr if tc.tag.endswith("}tc")]
                    out.append("| " + " | ".join(cells) + " |")
            out.append("")
    return "\n".join(out)

def main():
    src = Path(sys.argv[1]); dst = Path(sys.argv[2] if len(sys.argv) > 2 else "storyboard.md")
    ext = src.suffix.lower()
    if ext == ".docx": text = from_docx(src)
    elif ext == ".pdf":
        import pdfplumber
        with pdfplumber.open(src) as pdf: text = "\n".join(pg.extract_text() or "" for pg in pdf.pages)
    else: text = src.read_text(encoding="utf-8", errors="replace")
    dst.write_text(text, encoding="utf-8")
    n = text.count("## ") ; print(f"{src.name} → {dst} ({len(text)}자, 제목 줄 {n}개). 내용을 열어 장 구분·박스 칸이 살아 있는지 확인할 것")

if __name__ == "__main__":
    main()
