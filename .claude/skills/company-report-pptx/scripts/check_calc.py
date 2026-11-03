"""수식 검산: 원문에 적힌 계산 결과를 다시 계산해 맞는지 확인 (회로 설계 검토용)
사용법: python check_calc.py 검산.json
검산.json 예:
[{"id":"E2", "expr":"sqrt(2)*230/(10+0.5)", "stated":31, "unit":"A", "tol":0.05, "where":"p.3 필기 2번"}]
- expr: 파이썬 수식(math 함수 사용 가능: sqrt, pi, log, log10, exp, sin, cos ...). 단위는 SI로 맞춰 넣는다(mH→1e-3 등)
- tol: 허용 상대오차(기본 0.05 = 5%). 원문의 반올림 수준에 맞춘다
불일치가 있으면 종료코드 1 → 보고서에 [검산 불일치]로 표시하고 체크포인트 질문에 올린다
"""
import sys as _sys
for _s in (_sys.stdout, _sys.stderr):
    try: _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception: pass

import json, math, sys
SAFE = {k: getattr(math, k) for k in dir(math) if not k.startswith("_")}
def main():
    items = json.load(open(sys.argv[1], encoding="utf-8")); bad = 0
    for it in items:
        try: v = eval(it["expr"], {"__builtins__": {}}, SAFE)
        except Exception as e: print(f"[{it['id']}] 계산 오류: {e}"); bad += 1; continue
        st = it.get("stated"); tol = it.get("tol", 0.05); u = it.get("unit", "")
        if st is None: print(f"[{it['id']}] 계산값 {v:.4g} {u} (원문 결과 없음)"); continue
        rel = abs(v - st) / abs(st) if st else abs(v)
        ok = rel <= tol; bad += (not ok)
        print(f"[{it['id']}] {'일치' if ok else '불일치'}: 원문 {st} {u} / 재계산 {v:.4g} {u} (차이 {rel*100:.1f}%) {it.get('where','')}")
    print(f"검산 불일치 {bad}건"); sys.exit(1 if bad else 0)
if __name__ == "__main__": main()
