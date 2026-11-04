"""A3 장(PDCA 8칸) 내용 검사 — 칸이 '채워진 척'하는 것을 막는다.
사용법: python check_a3.py 결과.pptx
D.a3()/D.a3Slide()로 만든 칸(도형 이름 a3:<칸>)을 읽는다. A3 장이 없으면 바로 통과.
[실패] (종료코드 1) — 칸에 글은 있는데 그 칸이 답해야 할 것이 빠짐. 고치거나, 모르면 글을 비워 [확인 필요]로 드러낸다.
 - ③ 목표: 숫자 없음 또는 기한 없음
 - ⑥ 실행 계획: 담당 또는 기한 없음
 - ⑦ 효과 확인: 측정 지표(숫자)도, 측정 일정도 없음
[경고]
 - [확인 필요]·빈칸(___, ○월) 칸 목록 — 보고 전에 채울 곳
 - ① 문제 정의에 숫자(건수·시점·금액) 없음
 - ④ 근본 원인에 확신도 표시([확인됨]/[추정]/검증) 없음
 - ⑤ 대책에 임시/영구 구분 없음
"""
import sys as _sys
for _s in (_sys.stdout, _sys.stderr):
    try: _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception: pass
import sys, re
from pptx import Presentation

LABEL = {"problem": "① 문제 정의", "current": "② 현상 파악", "goal": "③ 목표", "cause": "④ 근본 원인",
         "action": "⑤ 대책", "plan": "⑥ 실행 계획", "check": "⑦ 효과 확인", "standard": "⑧ 표준화·횡전개"}
NUM = re.compile(r"\d")
DATE = re.compile(r"\d{1,2}/\d{1,2}|\d{1,2}\s*월|○\s*월|\d{4}\.\s*\d|\d+\s*주|까지|[1-4]\s*Q|Q[1-4]|분기|상반기|하반기|\d{1,2}\s*일")
OWNER = re.compile(r"담당|파트|그룹|팀|___|[가-힣]{2,4}\s*(프로|책임|선임|수석|님)")
BLANK = re.compile(r"확인 필요|___|○\s*월|○\s*일")

def main():
    prs = Presentation(sys.argv[1]); fails = warns = 0; found = False
    for n, sl in enumerate(prs.slides, 1):
        cells = {}
        for sh in sl.shapes:
            if (sh.name or "").startswith("a3:") and sh.has_text_frame:
                cells[sh.name[3:]] = sh.text_frame.text
        if not cells: continue
        found = True
        def F(k, why):
            nonlocal fails; fails += 1; print(f"[{n}장 {LABEL[k]}] 실패: {why} → 채우거나, 모르면 글을 비워 [확인 필요]로 드러낼 것")
        def W(k, why):
            nonlocal warns; warns += 1; print(f"[{n}장 {LABEL[k]}] 경고: {why}")
        blanks = [LABEL[k] for k, t in cells.items() if BLANK.search(t)]
        if blanks: warns += 1; print(f"[{n}장] 경고: 보고 전 채울 칸 — " + ", ".join(blanks))
        for k, t in cells.items():
            if t.startswith("[확인 필요]"): continue  # 미작성 칸은 의도적으로 드러낸 것
            if k == "goal":
                if not NUM.search(t): F(k, "목표 수치 없음 (예: 오트립 클레임 0건, 불량률 0.3% 이하)")
                if not DATE.search(t): F(k, "목표 기한 없음")
            elif k == "plan":
                if not OWNER.search(t): F(k, "담당 없음")
                if not DATE.search(t): F(k, "기한 없음")
            elif k == "check":
                if not NUM.search(t) and not ("측정" in t and DATE.search(t)): F(k, "측정 지표(숫자)도 측정 일정도 없음")
            elif k == "problem" and not NUM.search(t): W(k, "건수·시점·금액 등 숫자 없음")
            elif k == "cause" and not re.search(r"확인됨|추정|검증|재현|확인 필요", t): W(k, "원인 확신도 표시 없음 ([확인됨]/[추정])")
            elif k == "action" and not re.search(r"임시|영구|잠정|근본", t): W(k, "임시/영구 대책 구분 없음")
    if not found: print("A3 장 없음 — 검사 생략"); sys.exit(0)
    print(f"A3 이슈: 실패 {fails}건, 경고 {warns}건"); sys.exit(1 if fails else 0)

if __name__ == "__main__":
    main()
