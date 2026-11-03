"""결과 PPT 문구 검사: AI 티 표현, 번역투, 글자 예산, [확인 필요] 잔존.
사용법: python check_text.py <결과.pptx> [--extra 금지어1 금지어2 ...]
금지어 목록은 references/writing-style.md 의 '금지어' 절과 맞춰 관리한다.
"""
import sys as _sys
for _s in (_sys.stdout, _sys.stderr):
    try: _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception: pass

import sys, re, argparse
from pptx import Presentation

BANNED = ["다양한", "효과적으로", "을 통해", "를 통해", "다음과 같습니다", "혁신적", "중요한 역할",
          "시사점을 제공", "에 있어서", "하는 것이 필요합니다", "할 수 있습니다", "것으로 보입니다"]
MSG_MAX = 45  # 핵심 메시지 글자 예산

def main():
    ap = argparse.ArgumentParser(); ap.add_argument("pptx"); ap.add_argument("--extra", nargs="*", default=[])
    a = ap.parse_args(); banned = BANNED + a.extra
    prs = Presentation(a.pptx); issues = 0
    for i, sl in enumerate(prs.slides, 1):
        texts = [sh.text_frame.text for sh in sl.shapes if sh.has_text_frame and sh.text_frame.text.strip()]
        # 핵심 메시지 = 위에서 두 번째 큰 텍스트(헤더 규칙상 y≈1.0in)
        for sh in sl.shapes:
            if sh.has_text_frame and abs(sh.top / 914400 - 1.0) < 0.05 and len(sh.text_frame.text) > MSG_MAX:
                print(f"[{i}장] 핵심 메시지 {len(sh.text_frame.text)}자 > {MSG_MAX}자: {sh.text_frame.text}"); issues += 1
        for t in texts:
            for b in banned:
                if b in t: print(f"[{i}장] 금지 표현 '{b}': {t[:50]}"); issues += 1
            if "확인 필요" in t: print(f"[{i}장] [확인 필요] 남아 있음 → 사람 확인 후 제거: {t[:50]}")
    print(f"문구 이슈 {issues}건"); sys.exit(1 if issues else 0)

if __name__ == "__main__":
    main()
