# 테스트 시나리오: AS 무상지원금 – 일시불구독 개선 (ZCRMR058_1 / ZCRMR060_1 / RFC)

## 대상
- 프로그램: ZCRMR058_1, ZCRMR060_1, RFC `ZCRMF03_RFC_ZCRM_CHRG_CNFRM`, 테이블 `ZCRM_CHRG_CNFRM`
- T-code: 구독 ZCRM058S / ZCRM059S / ZCRM060S, 일반 ZCRM058A / ZCRM059A / ZCRM060A
- 대상 시스템/클라이언트: KDQ / 210 (KDP 금지)
- 원 시나리오: 김지인(G-ERP팀), 2026-09-11

## KDQ 사전 확인 결과 (2026-09-11, 읽기 전용)
| 항목 | 결과 |
|---|---|
| T-code 6종 | 모두 존재. 058A/058S/059A/059S → `ZCRMR058_1` 1000, 060A/060S → `ZCRMR060_1` 1000 |
| 058 선택화면 | `radPA_TAB` 확정 데이타 / `radPA_NEW` 전체(기본), `txtPA_SPMON` 기간월, `ctxtSO_VKBUR-*` 센터, `ctxtSO_KUNNR-*` 대리점, `ctxtSO_WEEK-*` 지급기준 |
| 060 선택화면 | `radPA_DISP` 조회(기본) / `radPA_TRAN` 전송, `txtSO_SPMON-*`, `ctxtSO_VKBUR-*`, `ctxtSO_KUNNR-*`, `ctxtSO_WEEK-*` (pa_subs 는 NO-DISPLAY 라 안 보임) |
| 선택화면 차이 | 058A/058S/059A/059S 선택화면·툴바가 **완전히 동일** (차이는 결과화면에서만 확인 가능) |
| RFC | SE37 표시 화면(SAPLSFUNCTION_BUILDER 3000) 진입 → 존재 |
| ZCRM_CHRG_CNFRM | 존재. 필드: 상담번호 / 월 / 대리점확정일 / 생성·변경 이력. **현재 0건** (테스트 전 정상 상태) |
| ZCRM_CICSVC_NEW-SUBS_TYPE | 존재 (SE16N 설명 "구독여부 지시자") |

### ZCRM_CICSVC_NEW 필드 매핑 (SE16N 설명 → 기술명)
| SE16N 설명 | 필드 | 비고 |
|---|---|---|
| 상담번호 | CSLTNO | |
| 상담번호의 seq(행)번호 | CSLTNO_ITEM | |
| 유무상구분자 | PRICETYPE | N=무상 |
| 조치일자 | RSLTDATE | 날짜형(DATS), SPMON 파생 원천 |
| 클레임유형 | CLMART | A/B/D1/D2/E1/E2 등 |
| 제외사유 | DRRSN | |
| 기사명 | TECHNAME | 기사코드(0000130211 형태) |
| 서비스점 | **KUNNR_SVC** | **대리점 = KUNNR_SVC.** `KUNNR` 는 고객코드. ZCRM_ASSPTFEEH1-KUNNR 값과 일치 |
| 영업소 | VKBUR | CT01~ 센터 |
| 구독여부 지시자 | SUBS_TYPE | |

- KUNNR_SVC 에 `CT01` 같은 센터 코드가 들어간 건이 있다 → 대리점 후보에서 제외.
- 월 데이터는 다음 달 10일경 적재(ERDAT 2026.02.10 = 2026.01 조치분).

## 테스트 데이터 선정 기준
- ZCRM_CICSVC_NEW: RSLTDATE = 대상월, PRICETYPE='N', CSLTNO_ITEM=1, CLMART ∈ {A,B,C,D,D1,D2}, DRRSN 공백, TECHNAME 있음, SUBS_TYPE 공백
- KUNNR-X: 대상월에 조건 충족 3건 이상인 대리점 → 3건을 'S' 로 변경
- KUNNR-Y: 다른 대리점 1건을 'S' 로 변경
- ZCRM_ASSPTFEEH1: 대상 SPMON + KUNNR-X/Y 에 행이 없어야 함 (확정/전표 이력 없음)
- ZCRM_CHRG_CNFRM: 비어 있음 (확인됨)

## 선정된 테스트 데이터 (2026-09-11, 기간월 2026.07)
| 역할 | 대리점(KUNNR_SVC) | 센터 | LIFNR | CSLTNO | RSLTDATE | CLMART |
|---|---|---|---|---|---|---|
| X-1 | 117786 제주진영서비스점 | CT11 | 120252 | 2026070100278 | 2026.07.01 | D1 |
| X-2 | 117786 | CT11 | 120252 | 2026070100450 | 2026.07.02 | D1 |
| X-3 | 117786 | CT11 | 120252 | 2026070100862 | 2026.07.01 | D1 |
| X-예비 | 117786 | CT11 | 120252 | 2026070101452 | 2026.07.01 | D1 |
| Y-1 | 100011 (주)승보상사 | CT02 | 123011 | 2026070102721 | 2026.07.02 | A |

- 모두 CSLTNO_ITEM=1, PRICETYPE=N, DRRSN 공백, TECHNAME 있음, SUBS_TYPE 공백(원복값 = 공백).
- KNVV VKORG=ASKR 존재, ZCRM_ASSPTFEEH1 2026.07 행 없음(2026.06 까지 전표 완료).
- **100001 제외**: 2026.07 FEEH1 행이 2026-09-11 에 이미 확정됨(타 테스트 추정).
- 058 실행 시 센터(SO_VKBUR) 필수.

## 확인 필요 / 리스크
1. **ZCRM_ASSPTFEEH1 키(SPMON+KUNNR)를 구독·일반이 공유**한다. 같은 대리점·같은 월에 일반 건이 058A 로 확정/전표되면 058S 전표(4-3)와 충돌 가능 → 개발자 확인 필요.
2. SUBS_TYPE 'S' 변경 방법: GUI 스크립팅으로 ABAP UPDATE 프로그램을 만들 수 없고, SE16N 편집(`chkGD-EDIT`)은 CLAUDE.md 규칙상 금지. → 개발자가 UPDATE 실행 또는 사용자 결정 필요.
3. 4-3(FIPOS)은 QA 에 실제 회계전표를 생성한다 → 실행 전 별도 승인.
4. SE16N `chkGD-EDIT` 가 매 실행 시 ON 으로 되돌아와 있음 (스크립트가 매번 끔).

## 테스트 케이스
원 시나리오 [그룹 1]~[그룹 5] 를 그대로 따른다. DB 변경 여부:

| 그룹 | 케이스 | DB 변경 |
|---|---|---|
| 0 | SUBS_TYPE 'S' 세팅 | **Y** (ZCRM_CICSVC_NEW) |
| 1 | 1-1 ~ 1-5 화면 분리 | N |
| 2 | 2-1, 2-2, 2-9 | **Y** (ZCRM_CHRG_CNFRM) |
| 2 | 2-3 ~ 2-8 | N (오류 기대, 단 2-5·2-6·2-7 은 정상건 동봉 시 Y) |
| 3 | 3-1 ~ 3-5 | 확정 등록/취소는 RFC 로 Y, 화면 확인은 N |
| 4 | 4-1, 4-2 | N |
| 4 | 4-3 전표생성 | **Y (회계전표)** |
| 4 | 4-4 미확정 전표 차단 | 오류 기대 (차단 실패 시 전표 생성 위험) |
| 5 | 5-1, 5-2 | RFC 오류 기대 / 화면 확인 |
| 5 | 5-3 원복 | **Y** |

## 결과 기록
- 실행일 / 로그 파일 / PASS·FAIL 요약: (진행 중)
