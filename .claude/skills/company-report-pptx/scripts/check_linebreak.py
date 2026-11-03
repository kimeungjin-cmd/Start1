"""한국어 줄바꿈·글자 넘침 검사 (글자 폭을 추정해 PowerPoint의 어절 단위 줄바꿈을 흉내 냄)
사용법: python check_linebreak.py <결과.pptx>
잡아내는 것
 1. 넘침: 예상 줄 수 × 줄 높이가 상자 높이를 넘음 → 글자가 상자 밖으로 나감
 2. 외톨이 줄: 자동 줄바꿈으로 마지막 줄에 1어절(또는 4글자 이하)만 남음 → 문장을 줄이거나 의미 단위로 직접 끊기
 3. 어색한 줄 시작: 줄이 조사·닫는 괄호·쉼표·단위로 시작
 4. 숫자와 단위가 갈림: '16' 다음 줄에 '시', '30' 다음 줄에 '분' 등
추정치라 렌더링 확인과 함께 쓴다. 표 안 글자는 검사하지 않는다.
"""
import sys as _sys
for _s in (_sys.stdout, _sys.stderr):
    try: _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception: pass

import sys, re
from pptx import Presentation
EMU = 914400
JOSA = ("은 는 이 가 을 를 의 에 로 와 과 도 만 께 서 이며 이고 이다 으로 에서 까지 부터").split()
BAD_START = tuple(JOSA) + (")", "]", ",", ".", "·", "%")  # "→ 결론" 형태의 줄 시작은 허용
UNITS = ("시", "분", "초", "건", "개", "장", "명", "원", "만원", "억", "년", "월", "일", "도", "V", "A", "mA", "mm", "㎜", "kg", "GB", "TB", "MHz", "%")

def cw(ch, pt):  # 글자 폭(인치)
    if ch == " ": return pt * 0.30 / 72
    if ord(ch) > 0x2E80: return pt * 0.98 / 72
    if ch.isupper() or ch.isdigit(): return pt * 0.60 / 72
    return pt * 0.50 / 72

def wrap(text, pt, width):
    lines, cur, curw = [], "", 0.0
    for word in text.split(" "):
        ww = sum(cw(c, pt) for c in word); sp = cw(" ", pt) if cur else 0
        if cur and curw + sp + ww > width:
            lines.append(cur); cur, curw = word, ww
        else:
            cur = (cur + " " + word) if cur else word; curw += sp + ww
    lines.append(cur); return lines

def main():
    prs = Presentation(sys.argv[1]); n_issue = 0
    for n, sl in enumerate(prs.slides, 1):
        for sh in sl.shapes:
            if not sh.has_text_frame or not sh.text_frame.text.strip() or sh.width is None: continue
            tf = sh.text_frame; W = sh.width / EMU; H = sh.height / EMU
            ml = (tf.margin_left or 0) / EMU; mr = (tf.margin_right or 0) / EMU
            width = max(W - ml - mr, 0.1); total_h = 0.0; out = []
            for p in tf.paragraphs:
                txt = "".join(r.text for r in p.runs)
                if not txt: continue
                sizes = [r.font.size.pt for r in p.runs if r.font.size] or [12]
                pt = max(sizes) * (1.04 if any(r.font.bold for r in p.runs) else 1)
                for seg in txt.split("\n"):
                    ls = wrap(seg, pt, width); out.append((seg, ls)); total_h += len(ls) * pt * 1.2 / 72
            short = tf.text.replace("\n", " / ")[:30]
            if total_h > H + 0.08:
                print(f"[{n}장] 넘침 예상 ({total_h:.2f}in > 상자 {H:.2f}in): {short}"); n_issue += 1
            for seg, ls in out:
                if len(ls) > 1 and (len(ls[-1].split(" ")) == 1 or len(ls[-1].replace(" ", "")) <= 4):
                    print(f"[{n}장] 외톨이 줄 '{ls[-1]}': {seg[:30]}"); n_issue += 1
            raw = tf.text.split("\n")
            for a, b in zip(raw, raw[1:]):
                b2 = b.strip()
                if b2.startswith(BAD_START):
                    print(f"[{n}장] 줄이 '{b2[:3]}'로 시작: {b2[:25]}"); n_issue += 1
                if re.search(r"\d$", a.strip()) and b2.startswith(UNITS):
                    print(f"[{n}장] 숫자·단위 분리: '{a.strip()[-6:]}' / '{b2[:6]}'"); n_issue += 1
    print(f"줄바꿈 이슈 {n_issue}건"); sys.exit(1 if n_issue else 0)

if __name__ == "__main__":
    main()
