# 조사 메모 — 전기오븐 7세그 LED 랜덤 점등(깜빡임)

원본 자료 없음(현상 설명만). 아래는 일반 기술 자료 기반. 시장 불량 수·모델·회로는 [확인 필요].

| # | 출처 | 핵심 |
|---|---|---|
| 1 | Nuvoton 기술 블로그 "Prevent ghosting on seven-segment display" https://forum.nuvoton.com/en/experience-sharing/tech-blog/14518 | 스캔 라인을 세그 데이터보다 먼저 바꾸면 고스트. 자리 전환 전 출력을 잠깐 끄는 블랭킹 필요 |
| 2 | Arduino forum "4 x 7 segment display ghosting" https://forum.arduino.cc/t/4-x-7-segment-display-ghosting/439058 | 모든 자리 OFF → 세그 데이터 출력 → 새 자리 ON 순서 |
| 3 | AllAboutCircuits "7 segment multiplex ghosts" https://forum.allaboutcircuits.com/threads/7-segment-multiplex-ghosts.115736/ | PNP 하이사이드 스위치 턴오프 지연(B-E 용량)이 고스트 원인 |
| 4 | Titan Micro TM1628A 데이터시트(LCSC C2943005) https://www.lcsc.com/product-detail/C2943005.html | 전원 투입 시 표시 레지스터 임의값 → C0H~CDH에 0x00 14바이트 기록 후 표시 ON 권고 |
| 5 | Arduino forum "MAX7219 display errors" https://forum.arduino.cc/t/max7219-display-errors/243195 | DATA/CLK 글리치로 드라이버 레지스터 오염 → 재초기화 전까지 오표시 유지 |
| 6 | Arduino forum "EMI from relay" https://forum.arduino.cc/t/emi-from-relay/333008 외 | 릴레이·접촉기 개폐 노이즈 → MCU 리셋·표시 오동작. 스너버, 전원 디커플링, 배선 분리 |
| 7 | EN 55014-2(CISPR 14-2) 요약, testxchange https://www.testxchange.com/pruefnorm/din-en-55014-2 | 가전 내성 ESD: 접촉 4 kV / 기중 8 kV (IEC 61000-4-2) |
| 8 | LED 디스플레이 데이터시트(Broadcom AV02-1644EN 등) https://docs.broadcom.com/doc/AV02-1644EN | 역바이어스 → 금속 이동·리크 증가, 고습 급격한 온도 변화 시 결로 주의 |

미확인: EN 55014-2 EFT 시험 레벨(AC 전원·신호 포트) → 사내 시험 규격으로 확인 필요.
