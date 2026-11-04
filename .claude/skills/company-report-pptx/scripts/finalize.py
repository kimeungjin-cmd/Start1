"""완료 게이트: 모든 검사 + 렌더링 + '장별 육안 검수 기록'이 있어야만 통과.
사용법: python finalize.py 결과.pptx [원본파일 ...] [--ignore 숫자 ...]
1회차: 검사 4종 실행 → 렌더링(render/) → review.md 양식 생성 → 종료코드 2 (아직 미완료)
       → AI는 render/slide-XX.png 를 '한 장씩 직접 열어 보고' review.md 를 채운다
2회차: 검사 재실행 + review.md 검증(장 수·항목·렌더링 토큰 일치) → 통과 시 종료코드 0, "FINALIZE PASS"
PPT를 고치면 렌더링 토큰이 바뀐다. 이때 장별 이미지가 이전과 '픽셀까지 같은' 장은 이전 검수 기록을 자동으로 이어받고,
바뀐 장만 다시 검수하면 된다(토큰 절약). 이미지 해시로 판별하므로 바뀐 장을 건너뛸 수는 없다.
"""
import sys as _sys
for _s in (_sys.stdout, _sys.stderr):
    try: _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception: pass

import sys, os, re, subprocess, hashlib, argparse, glob
from pathlib import Path
HERE = Path(__file__).parent
FIELDS = ["본 이미지", "넘침·잘림", "줄바꿈", "정렬", "10초 질문", "10초 답이 보이나", "판정"]

def run(cmd):
    r = subprocess.run([sys.executable] + cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", env={**os.environ, "PYTHONIOENCODING": "utf-8"})
    print(r.stdout.strip().splitlines()[-1] if r.stdout.strip() else r.stderr.strip()[-200:])
    return r.returncode == 0, r.stdout

def token(pptx):
    return hashlib.sha1(Path(pptx).read_bytes()).hexdigest()[:10]

def img_hash(p):
    return hashlib.md5(Path(p).read_bytes()).hexdigest()[:12]

def blocks_of(text):
    parts = re.split(r"^## (\d+)장\s*$", text, flags=re.M)
    return {int(parts[i]): parts[i + 1] for i in range(1, len(parts) - 1, 2)}

def template(n, tok, hashes, prev_text=""):
    prev = blocks_of(prev_text) if prev_text else {}
    out = [f"# 장별 육안 검수 기록\n렌더링 토큰: {tok}\n",
           "finalize가 알려준 모아보기(todo-sheet-N.jpg) 또는 review/slide-XX.jpg를 한 번에 한 장씩 직접 열어 본 뒤 채운다. 빈칸·'TODO'가 있으면 통과하지 않는다.",
           "판정은 'OK' 또는 '수정 후 OK'. 수정했다면 PPT 재생성 → finalize 재실행(토큰이 바뀌므로 다시 검수).\n"]
    reused = 0
    for i in range(1, n + 1):
        h = hashes[i]; b = prev.get(i, "")
        m = re.search(r"^- 이미지 해시:\s*(\w+)", b, flags=re.M); j = re.search(r"^- 판정:\s*(.*)$", b, flags=re.M)
        if m and m.group(1) == h and j and j.group(1).strip() in ("OK", "수정 후 OK") and "TODO" not in b:
            out.append(f"## {i}장"); out.append(b.strip("\n").replace("(이전 검수 유지 — 이미지 동일)", "").strip("\n")); out.append("(이전 검수 유지 — 이미지 동일)\n"); reused += 1
        else:
            out.append(f"## {i}장"); out.append(f"- 이미지 해시: {h}")
            out += [f"- {f}: TODO" for f in FIELDS]; out.append("")
    return "\n".join(out), reused

def validate(review, n, tok, hashes):
    t = Path(review).read_text(encoding="utf-8"); errs = []
    if f"렌더링 토큰: {tok}" not in t: errs.append("렌더링 토큰 불일치 — PPT가 바뀐 뒤 다시 검수하지 않았음")
    blocks = re.split(r"^## (\d+)장\s*$", t, flags=re.M)
    got = {int(blocks[i]): blocks[i + 1] for i in range(1, len(blocks) - 1, 2)}
    for i in range(1, n + 1):
        b = got.get(i)
        if b is None: errs.append(f"{i}장 검수 기록 없음"); continue
        m = re.search(r"^- 이미지 해시:\s*(\w+)", b, flags=re.M)
        if not m or m.group(1) != hashes[i]: errs.append(f"{i}장 이미지 해시 불일치 — 바뀐 장을 다시 검수해야 함")
        for f in FIELDS:
            m = re.search(rf"^- {re.escape(f)}:\s*(.*)$", b, flags=re.M)
            v = (m.group(1).strip() if m else "")
            if not v or "TODO" in v: errs.append(f"{i}장 '{f}' 비어 있음")
        img = re.search(r"^- 본 이미지:\s*(.*)$", b, flags=re.M)
        if img and f"slide-{i:02d}" not in img.group(1): errs.append(f"{i}장 '본 이미지'에 slide-{i:02d}.png가 없음")
        j = re.search(r"^- 판정:\s*(.*)$", b, flags=re.M)
        if j and j.group(1).strip() not in ("OK", "수정 후 OK"): errs.append(f"{i}장 판정이 OK가 아님: {j.group(1).strip()}")
    return errs

def main():
    ap = argparse.ArgumentParser(); ap.add_argument("pptx"); ap.add_argument("sources", nargs="*")
    ap.add_argument("--ignore", nargs="*", default=[]); a = ap.parse_args()
    from pptx import Presentation
    n = len(Presentation(a.pptx).slides); tok = token(a.pptx)
    work = Path(a.pptx).resolve().parent; rdir = work / "render"; review = work / "review.md"
    print("== 자동 검사"); ok = True
    for name in ["check_layout.py", "check_linebreak.py", "check_text.py", "check_richness.py", "check_a3.py"]:
        ok &= run([str(HERE / name), a.pptx])[0]
    if a.sources:
        ok &= run([str(HERE / "check_coverage.py"), a.pptx] + a.sources + (["--ignore"] + a.ignore if a.ignore else []))[0]
    stamp = rdir / ".token"
    if not (stamp.exists() and stamp.read_text() == tok):
        print("== 렌더링"); subprocess.run([sys.executable, str(HERE / "render_pptx.py"), a.pptx, str(rdir)], check=True, env={**os.environ, "PYTHONIOENCODING": "utf-8"})
        stamp.write_text(tok)
    hashes = {i: img_hash(rdir / f"slide-{i:02d}.png") for i in range(1, n + 1)}
    if not review.exists() or f"렌더링 토큰: {tok}" not in review.read_text(encoding="utf-8"):
        prev = review.read_text(encoding="utf-8") if review.exists() else ""
        body, reused = template(n, tok, hashes, prev)
        review.write_text(body, encoding="utf-8")
        todo = [i for i in range(1, n + 1) if f"## {i}장\n- 이미지 해시: {hashes[i]}\n- 본 이미지: TODO" in body]
        if not todo:
            print(f"\n모든 장이 이전 검수와 이미지 동일 → 기록 이어받음 ({reused}장)")
        else:
            sys.path.insert(0, str(HERE)); import render_pptx
            for old in rdir.glob("todo-sheet-*.jpg"): old.unlink()
            sh = render_pptx.sheets(str(rdir), todo, prefix="todo-sheet")
            print(f"\n[미완료] 다시 볼 장 {len(todo)}개 (이미지 동일로 이어받은 장 {reused}개)")
            print("  아래 모아보기 이미지를 **한 번에 한 장씩** 열어 본다(여러 이미지를 한 요청에 넣지 말 것 — 입력 초과 오류 방지):")
            for fpath, grp in sh: print(f"   - {fpath}  →  " + ", ".join(f"slide-{i:02d}" for i in grp))
            print(f"  글자가 작아 판단이 어려운 장만 {rdir / 'review'}/slide-XX.jpg 를 따로 연다.")
            print(f"  그다음 {review} 의 TODO를 채우고 finalize를 다시 실행한다. '본 이미지' 칸엔 연 파일명과 slide 번호를 적는다(예: todo-sheet-1.jpg slide-02).")
            sys.exit(2)
    errs = validate(review, n, tok, hashes)
    if errs or not ok:
        print("\n[미완료]"); [print(" -", e) for e in errs]
        if not ok: print(" - 자동 검사 이슈가 남아 있음 (위 출력 참고)")
        sys.exit(1)
    print(f"\nFINALIZE PASS — 자동 검사 통과, {n}장 육안 검수 기록 확인 (토큰 {tok})")

if __name__ == "__main__":
    main()
