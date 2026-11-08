// 전기오븐 7세그 LED 랜덤 점등(깜빡임) — 예상원인과 대책
const path = require("path");
const D = require(path.join(__dirname, "../../.claude/skills/company-report-pptx/scripts/deck.js"));
const pres = D.createDeck(), F = f => path.join(__dirname, f);
const SRC_GEN = "※ 근거: 드라이버 IC 데이터시트·고스트 대책 기술자료(유첨 출처) / 시장 불량 수·대상 모델은 미입수";

// 1 한 장 요약 [맞춤]
(() => {
  const { slide, boxes } = D.slide(pres, {
    title: "전기오븐 LED 표시 랜덤 점등 불량 원인과 대책",
    message: "예상원인 4계열 중 '통신선 노이즈로 표시 데이터 오염'이 1순위 [추정]",
    sub: "회수품 분석 전, 증상 설명 기준 판단 / 재현시험으로 원인 확정 후 HW·FW 대책 병행",
    layout: "2-64", source: SRC_GEN,
    boxes: [{ title: "예상원인 4계열", msg: "랜덤·순간 점등 증상은 A계열과 가장 부합" },
            { title: "원인 확정과 조치 순서", msg: "증상 패턴 확인이 첫 분기점" }],
    notes: "원본 없음(현상 설명만). 1순위 판단 근거: '랜덤', '깜빡임처럼 보임' = 다음 갱신 때 사라지는 순간 오점등 → 데이터 오염 계열. 회수품·발생 조건 확인 시 바뀔 수 있음",
  });
  D.cards(pres, slide, boxes[0].inner, [
    { tag: "A · 1순위", red: true, title: "통신선 노이즈 → 표시 데이터 오염", body: "히터 릴레이·모터 개폐 순간 엉뚱한 세그먼트 점등" },
    { tag: "B", title: "전원 강하·정전기 → 드라이버 오동작", body: "표시 전체 깜빡임, 재부팅 전까지 이상 유지" },
    { tag: "C", title: "구동 타이밍 → 고스트(잔상)", body: "꺼진 세그먼트가 늘 같은 자리에서 희미하게 점등" },
    { tag: "D", title: "고온·습기·접촉 불량", body: "조리 중 고온, 스팀 후, 결로 때만 발생" },
  ]);
  D.steps(pres, slide, boxes[1].inner, [
    ["증상 패턴 확인", "발생 조건·밝기·지속 시간\n→ 원인 계열 압축", true],
    ["재현시험", "릴레이 동기 파형, ESD,\n온습도 시험으로 확정"],
    ["대책 적용", "FW 주기 재기록(즉시)\n+ HW 필터·블랭킹"],
    ["효과 검증", "재현 조건 반복 시험\n→ 횡전개 대상 확정"],
  ]);
})();

// 2 현상 정의와 구조
(() => {
  const { slide, boxes } = D.slide(pres, {
    title: "현상 정의와 표시부 구조",
    message: "MCU → 하네스 → 드라이버 IC → LED 경로 중 4곳이 원인 후보",
    sub: "7세그는 자리를 빠르게 번갈아 켜는 다이나믹 구동, 경로 어디서든 오점등 가능",
    layout: "2-55", source: "※ 구조는 일반적인 오븐 표시부 기준 / 실제 회로·드라이버 IC 품번은 [확인 필요]",
    boxes: [{ title: "표시부 신호 경로와 원인 위치", msg: "A~D는 다음 장 판별표와 대응" },
            { title: "현상 정의", msg: "꺼져야 할 세그먼트가 랜덤하게 점등" }],
    notes: "블록도: 메인 PCB(MCU, SMPS, 히터 릴레이·모터) - 하네스(CLK·DIO·STB, VDD·GND) - 디스플레이 PCB(드라이버 IC, 7세그 LED). 원인 위치 A 하네스 신호선, B 전원선, C 드라이버~LED 구동, D 디스플레이 PCB 환경",
  });
  const A = boxes[0].inner;
  D.figureFit(pres, slide, { x: A.x, y: A.y, w: A.w, h: 3.55 }, F("fig_block.png"), "diagram");
  D.rows(pres, slide, boxes[1].inner, [
    ["증상", "꺼진 세그먼트가 순간 켜졌다 꺼짐\n→ 사용자에게 깜빡임으로 보임", true],
    ["표시 방식", "다이나믹 구동(자리 순차 점등)\n드라이버 IC 직렬 통신 [확인 필요]"],
    ["미확인", "발생률·대상 모델·발생 조건\n→ 시장 VOC·회수품으로 확인"],
    ["판단 범위", "증상 설명과 일반 기술자료 기준\n회수품 분석 전 단계"],
  ], { lw: 1.15 });
})();

// 3 증상 패턴별 판별표 [맞춤 근거]
(() => {
  const { slide, boxes } = D.slide(pres, {
    title: "증상 패턴으로 원인 계열 판별",
    message: "밝기·지속 시간·발생 조건 3가지만 확인하면 4계열 중 1~2개로 압축",
    sub: "현장·회수품에서 먼저 묻고 볼 항목, 판별 후 재현시험 범위를 좁힘",
    layout: "1", source: "※ 판별 기준: 고스트·드라이버 레지스터 오염 사례 기술자료(유첨) 기반, 실측 검증 전 [추정]",
    boxes: [{ title: "증상 → 원인 → 판별 방법", msg: "'정상 밝기로 순간 점등 + 릴레이 동작과 동기'면 A" }],
    notes: "판별 3요소: 밝기(정상/희미), 지속(순간/계속/재부팅까지), 조건(릴레이·모터 동작, 정전기, 고온·스팀)",
  });
  D.table(pres, slide, boxes[0].inner,
    ["계열", "증상 패턴", "예상 원인", "판별 방법"],
    [
      ["A", "정상 밝기로 순간 점등, 히터·팬 ON/OFF 순간 발생", "CLK·DIO 노이즈로 잘못된 데이터 래치", "릴레이 동작과 CLK·DIO 파형 동기 측정"],
      ["B", "표시 전체 깜빡임·꺼짐, 재부팅까지 유지", "VDD 강하·ESD로 드라이버 리셋·오동작", "VDD 강하 측정, 표시부 ESD 인가"],
      ["C", "늘 같은 자리에 희미하게 점등, 어두우면 잘 보임", "자리 전환 시 블랭킹 부족, 턴오프 지연", "COM·SEG 타이밍 측정, '8888' 표시 비교"],
      ["D", "고온 조리 중·스팀 후·결로 시에만 발생", "LED·PCB 리크, 커넥터 간헐 접촉", "온습도 시험, 리크 전류·접촉 저항 측정"],
    ],
    { colW: [0.8, 3.9, 3.4, 3.83], hi: 0, rowH: 0.7, fontSize: 12 });
  D.boxNote(slide, boxes[0].inner, "※ A가 1순위인 이유: '랜덤'·'깜빡임처럼 보임'은 다음 화면 갱신 때 사라지는 순간 오점등 패턴과 일치");
})();

// 4 원인 A·B 노이즈·전원
(() => {
  const { slide, boxes } = D.slide(pres, {
    title: "원인 A·B: 개폐 노이즈와 전원 강하",
    message: "히터 릴레이 개폐 노이즈가 하네스로 들어가 표시 데이터를 오염 [추정]",
    sub: "드라이버 레지스터가 오염되면 다음 재기록 전까지 오표시 유지, 재기록 주기가 길수록 길게 보임",
    layout: "2-55", source: "※ 근거: 드라이버 레지스터 오염 사례, 릴레이 노이즈 대책 기술자료(유첨)",
    boxes: [{ title: "발생 메커니즘", msg: "노이즈 1회가 오점등 1회로 이어지는 사슬" },
            { title: "확인할 설계 포인트", msg: "필수 4개는 도면·실측으로 바로 확인" }],
    notes: "A: 릴레이·모터 개폐 → 하네스 유도 → CLK 오인식 → 오데이터 래치 → 오점등. B: VDD 강하·ESD → 드라이버 리셋·래치업 → 전체 깜빡임·이상 유지",
  });
  const [a, b] = D.splitV(boxes[0].inner, [1, 1], 0.2);
  D.flow(pres, slide, a, [
    { title: "A 릴레이 개폐", desc: "히터·팬", state: "risk" },
    { title: "하네스 유도", desc: "CLK·DIO", state: "now" },
    { title: "오점등", desc: "다음 갱신까지", state: "plan" },
  ]);
  D.flow(pres, slide, b, [
    { title: "B 전원 강하", desc: "부하 ON·ESD", state: "risk" },
    { title: "드라이버 오동작", desc: "리셋·래치업", state: "now" },
    { title: "전체 깜빡임", desc: "재부팅까지", state: "plan" },
  ]);
  D.checklist(pres, slide, boxes[1].inner, [
    ["표시 데이터 재기록 주기", "한 번 쓰고 끝인지, 주기적 전체 재기록인지", "필수"],
    ["CLK·DIO 필터", "직렬 저항·커패시터 유무, 하네스 길이", "필수"],
    ["드라이버 VDD 디커플링", "IC 핀 가까이 0.1µF 유무", "필수"],
    ["릴레이 스너버·배선 분리", "개폐 경로와 신호선 거리, GND 공통 구간", "필수"],
    ["ESD 경로", "표시 창·버튼 주변 방전 경로", "선택"],
  ]);
})();

// 5 원인 C 고스트 [맞춤]
(() => {
  const { slide, boxes } = D.slide(pres, {
    title: "원인 C: 자리 전환 타이밍(고스트)",
    message: "자리 전환 때 세그 데이터가 늦게 바뀌면 이웃 자리 패턴이 희미하게 점등",
    sub: "랜덤보다는 '늘 같은 자리·같은 패턴'이 특징, A와 증상으로 구분 가능",
    layout: "2-55", source: "※ 근거: Nuvoton 고스트 방지 기술자료, 7세그 멀티플렉싱 고스트 사례(유첨)",
    boxes: [{ title: "구동 타이밍 비교", msg: "블랭킹 구간 하나로 겹침 제거" },
            { title: "고스트를 만드는 요인", msg: "FW 순서와 HW 턴오프 속도 둘 다 확인" }],
    notes: "타이밍도: 블랭킹 없을 때 자리2 ON 후 세그 데이터 변경 → 겹침 구간에 자리1 패턴이 자리2에 표시. 대책: 모든 자리 OFF → 세그 데이터 변경 → 다음 자리 ON",
  });
  const A = boxes[0].inner;
  D.figureFit(pres, slide, { x: A.x, y: A.y, w: A.w, h: 3.55 }, F("fig_timing.png"), "diagram");
  D.rows(pres, slide, boxes[1].inner, [
    ["FW 순서", "자리 ON 후 세그 변경하는 순서\n→ 모든 자리 OFF 후 변경", true],
    ["턴오프 지연", "자리 구동 트랜지스터가 늦게 꺼짐\n(베이스 저항·포화 상태)"],
    ["기생 용량", "COM 라인에 남은 전하가\n다음 자리에서 방전"],
    ["드라이버 IC", "IC 내부 스캔이면 IC 설정 확인\n외부 스캔이면 위 3가지 확인"],
  ], { lw: 1.2 });
})();

// 6 원인 D 환경·접촉
(() => {
  const { slide, boxes } = D.slide(pres, {
    title: "원인 D: 고온·습기·접촉 불량",
    message: "오븐 특유의 고온·스팀 환경에서만 나면 리크·접촉 계열로 판단",
    sub: "발생 조건이 조리 중·조리 직후에 몰리는지가 판별 핵심",
    layout: "2-55", source: "※ 근거: LED 디스플레이 데이터시트(역바이어스·결로 주의), 유첨 출처",
    boxes: [{ title: "환경 기인 메커니즘", msg: "누설 경로가 생기면 꺼진 세그먼트에 미세 전류" },
            { title: "확인 항목", msg: "회수품 외관과 환경 시험으로 확인" }],
    notes: "LED 데이터시트: 역바이어스는 금속 이동·리크 증가 유발, 고습에서 급격한 온도 변화 시 결로 주의",
  });
  D.cards(pres, slide, boxes[0].inner, [
    { tag: "고온", title: "LED·IC 리크 증가", body: "도어·캐비티 열로 표시부 온도 상승" },
    { tag: "습기", red: true, title: "결로·플럭스 잔사 누설", body: "스팀 후 PCB 표면에 세그먼트 간 누설 경로" },
    { tag: "접촉", title: "커넥터·납땜 간헐 접촉", body: "열 사이클로 FFC 접점·납땜부 들뜸" },
  ]);
  D.checklist(pres, slide, boxes[1].inner, [
    ["회수품 PCB 외관", "부식·잔사·결로 흔적, 커넥터 접점", "필수"],
    ["표시부 동작 온도", "조리 중 표시부 실측 온도 vs 부품 정격", "필수"],
    ["고온 고습 재현", "스팀 조리 후 표시 상태 관찰", "필수"],
    ["코팅·세척 공정", "방습 코팅 유무, 세척 조건", "선택"],
  ]);
})();

// 7 대책안
(() => {
  const { slide, boxes } = D.slide(pres, {
    title: "대책안: FW 즉시 적용, HW는 원인 확정 후",
    message: "FW 주기 재기록은 원인 확정 전에도 A·B 증상 시간을 줄이는 1차 대책",
    sub: "HW 변경은 재현시험으로 원인 계열을 확정한 뒤 해당 대책만 적용",
    layout: "2-55", source: "※ 근거: TM1628A 데이터시트(초기화 권고), 고스트·노이즈 대책 기술자료 / 위치도는 해석값",
    boxes: [{ title: "대책별 효과와 적용 용이성", msg: "FW 재기록이 가장 빠르고 범위 넓음" },
            { title: "대책 목록", msg: "원인 계열별로 짝을 지어 적용" }],
    notes: "위치도 x=적용 용이성, y=효과 범위, 해석값(정답 아님). 비용·일정은 [확인 필요]",
  });
  D.matrix(pres, slide, boxes[0].inner, {
    xLabel: "적용 용이성", yLabel: "효과 범위", goodZone: "tr", goodText: "먼저 적용",
    items: [
      { name: "FW 주기 재기록", x: 0.82, y: 0.78, red: true, note: "즉시" },
      { name: "FW 블랭킹", x: 0.72, y: 0.35 },
      { name: "신호선 RC 필터", x: 0.4, y: 0.7 },
      { name: "릴레이 스너버", x: 0.22, y: 0.55 },
      { name: "방습 코팅", x: 0.15, y: 0.22 },
    ],
  });
  D.boxNote(slide, boxes[0].inner, "※ 위치는 해석값(정답 아님), 비용 반영 전");
  D.table(pres, slide, boxes[1].inner,
    ["대상", "대책", "구분"],
    [
      ["A·B", "표시 데이터 주기적 전체 재기록, 표시 ON 명령 재전송", "FW"],
      ["A", "CLK·DIO 직렬 저항+커패시터, 하네스 경로 분리", "HW"],
      ["A·B", "드라이버 VDD 디커플링 보강, 릴레이 스너버", "HW"],
      ["C", "모든 자리 OFF → 세그 변경 → 다음 자리 ON", "FW"],
      ["D", "방습 코팅, 세척 조건, 커넥터 사양 재검토", "공정"],
    ],
    { colW: [0.75, 3.95, 0.94], rowH: 0.6, fontSize: 11 });
})();

// 8 검증 계획
(() => {
  const { slide, boxes } = D.slide(pres, {
    title: "원인 확정과 대책 검증 계획",
    message: "재현시험 5종으로 원인 계열 확정 후 대책 효과를 같은 조건에서 재확인",
    sub: "일정·담당은 제안안, 회수품 입수 시점에 맞춰 확정 [확인 필요]",
    layout: "2-55", source: "※ ESD 레벨: EN 55014-2(CISPR 14-2) 요약 기준 / EFT 레벨은 사내 시험 규격 확인 [확인 필요]",
    boxes: [{ title: "진행 일정(제안)", msg: "증상 확인부터 횡전개까지 5단계" },
            { title: "재현시험 조건", msg: "원인 계열마다 시험 1개씩 대응" }],
    notes: "ESD: 접촉 4 kV / 기중 8 kV (IEC 61000-4-2 방법, EN 55014-2 요약). EFT 레벨 미확인. 일정 상대 주차는 제안값",
  });
  D.timeline(pres, slide, boxes[0].inner, [
    { date: "1주", title: "증상·회수품 확인" },
    { date: "2주", title: "재현시험", red: true },
    { date: "2주", title: "FW 재기록 적용" },
    { date: "3~4주", title: "HW 대책 시작품" },
    { date: "5주", title: "효과 검증", red: true },
  ]);
  D.boxNote(slide, boxes[0].inner, "※ 착수 기준 상대 주차, 담당·날짜 [확인 필요]");
  D.table(pres, slide, boxes[1].inner,
    ["계열", "시험", "조건"],
    [
      ["A", "릴레이 동기 파형 측정", "히터·팬 반복 개폐, CLK·DIO 관찰"],
      ["B", "ESD 인가", "접촉 4 kV / 기중 8 kV"],
      ["B", "전원 강하 측정", "전 부하 ON 시 드라이버 VDD"],
      ["C", "구동 타이밍 측정", "COM·SEG 동시 측정, 암실 촬영"],
      ["D", "고온 고습 재현", "스팀 조리 후 관찰 [확인 필요]"],
    ],
    { colW: [0.7, 2.0, 2.94], rowH: 0.6, fontSize: 11 });
})();

// 유첨 1 출처
(() => {
  const { slide, boxes } = D.slide(pres, {
    title: "유첨1. 참고 기술자료",
    message: "원인 판단은 일반 기술자료 기준, 사내 회로·회수품 확인으로 보완 필요",
    sub: "출처는 데이터시트·규격 요약·제조사 기술자료·엔지니어 포럼, 포럼 사례는 참고 수준",
    layout: "1", source: "※ 조사일 2026-11-08",
    boxes: [{ title: "출처와 활용 장", msg: "규격 수치는 ESD 레벨 1건만 확인, EFT는 미확인" }],
    notes: "research.md 참조",
  });
  D.table(pres, slide, boxes[0].inner,
    ["구분", "자료", "핵심", "활용"],
    [
      ["데이터시트", "Titan Micro TM1628A", "전원 투입 시 표시 레지스터 0x00 기록 후 표시 ON 권고", "7장"],
      ["데이터시트", "Broadcom LED 디스플레이 AV02-1644EN 외", "역바이어스 리크 증가, 고습 결로 주의", "6장"],
      ["규격 요약", "EN 55014-2 (CISPR 14-2)", "가전 ESD 접촉 4 kV / 기중 8 kV", "8장"],
      ["기술자료", "Nuvoton 고스트 방지 블로그", "자리 전환 전 출력 OFF(블랭킹)", "5장"],
      ["포럼 사례", "Arduino·AllAboutCircuits", "레지스터 오염·턴오프 지연·릴레이 노이즈", "4·5장"],
    ],
    { colW: [1.3, 3.4, 5.73, 1.5], rowH: 0.55, fontSize: 11 });
})();

D.save(pres, F("전기오븐_LED_7세그_랜덤점등_원인과대책.pptx"), { reserveCover: true });
