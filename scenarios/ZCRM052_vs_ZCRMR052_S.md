# 테스트 시나리오: ZCRM052(ZCRMR052) vs ZCRMR052_S 비교

## 대상
- 기존: T-code `ZCRM052` → 프로그램 `ZCRMR052`, 화면 제목 `회수품적합성관리`
- 신규: T-code `ZCRM052_S` → 프로그램 `ZCRMR052_S` (구독 전용)
  - `ZCRM052_S` 는 이 문서에서 쓰는 표기다. 실제 임시 T-code 값은 `scripts/config_local.txt` 에만 두고 문서에는 적지 않는다.
- 시스템/클라이언트: KDQ / 210 (KDP 금지). 사용자 KKR003-1, 언어 KO
- 조사일: 2026-09-09

## 목적
구독 구분(`ZCRM_CICSVC-SUBS_TYPE`) 로직을 제외한 조회·화면·검증·저장·상세조회 기능이 기존과 동일하게 동작하는지 확인한다.

## 비교 제외 항목
신규 프로그램의 `SUBS_TYPE` 컬럼 / 프로그램명 / 신규 T-code / 화면 제목 / ALV Layout Variant / 구독 데이터 필터 조건.

---

## 1. 화면 ID 참조 (2026-09-09 KDQ 210 실측, 두 프로그램 동일)

### 선택화면 (screen 1000) — `wnd[0]/usr/` 하위
| 항목 | ID | 비고 |
|---|---|---|
| 처리조직 | `radPA_R01` 센터조치 / `radPA_R02` 서비스점조치 | 라디오 변경 후 Enter 필요 |
| 센터 | `ctxtSO_VKBUR-LOW` / `-HIGH` | **PA_R01 일 때만 존재** |
| 서비스점 | `ctxtSO_KUNNR-LOW` / `-HIGH` | **PA_R02 일 때만 존재** |
| 회수품요청일자(청구요청일) | `ctxtSO_PRDRE-LOW` / `-HIGH` | 항상 존재 |
| 상담번호 | `txtSO_CSLTN-LOW` / `-HIGH` | |
| 접수일자 | `ctxtSO_RCDT-LOW` / `-HIGH` | **PA_CHX 일 때만 존재**, 기본값 오늘~오늘 |
| 제품 계층구조 | `ctxtSO_PRDH-LOW` / `-HIGH` | |
| 유/무상 | `ctxtSO_PRTY-LOW` | 단일값 |
| 자재번호(부품) | `txtSO_PRT-LOW` / `-HIGH` | |
| 멀티킷 사용 여부 | `ctxtSO_MULTK-LOW` | 단일값 |
| 접수여부 | `radPA_CHX` 회수품 접수 / `radPA_NCH` 회수품 미접수 | |
| 실행 | `wnd[0]/tbar[1]/btn[8]` (F8) | |

결과화면 사용자 버튼: `wnd[0]/tbar[1]/btn[23]` 회수가능 제조일 상세, `btn[26]` 부적합 사유, `btn[27]` 책임사업장. 두 프로그램 동일.

> 필드는 비활성(disabled)이 아니라 **화면에서 사라지는 방식**이다. 따라서 판정은
> `findById` 성공/실패(존재 여부)로 한다. `Changeable` 은 존재하는 필드에서 항상 True.

### 결과화면 (screen 100)
| 항목 | ID |
|---|---|
| 접수부서 | `wnd[0]/usr/txtGS_HEADER-REMARK1` |
| 접수자 | `wnd[0]/usr/txtGS_HEADER-REMARK2` |
| ALV 그리드 | `wnd[0]/usr/cntlCON1/shellcont/shell` (GuiShell / GridView) |
| 상세조회 버튼 | `wnd[0]/tbar[1]/btn[23]` |
| 조회 건수 | 화면 제목 `회수품적합성관리 (n)` 의 괄호 안 숫자 |

---

## 2. 사전 조건

1. **구독조건 임시 비활성화**: 신규 프로그램의 회수품 접수 SELECT / 미접수 SELECT 두 곳에서
   `AND zcrm_cicsvc~subs_type = gc_subs_type_s` 가 주석 처리돼 있어야 한다.
   → **2026-09-09 KDQ 에서 주석 처리 상태로 확인됨** (사용자 확인 + 조회 건수 일치로 교차확인).
   테스트 종료 후 접수/미접수 SELECT 양쪽 모두 원복해야 한다.
2. 두 프로그램을 같은 시스템·클라이언트·사용자로 실행한다.
3. 저장 테스트(TC-07·08·13·16)는 기존/신규에 **서로 다른 데이터**를 쓴다.
   같은 데이터를 쓸 경우 기존 프로그램 테스트 후 DB 를 원복해야 한다.
4. 저장 테스트 시 접수부서/접수자에 `ZTEST_20260909` 식별자를 넣어 사후 정리 가능하게 한다.

---

## 3. 테스트 케이스

### 1차 비저장 비교 (DB 변경 없음)

| ID | 내용 | 입력 | 기대 결과 | DB |
|---|---|---|---|---|
| TC-01 | 센터조치 + 회수품 접수 | `PA_R01`+`PA_CHX`, `SO_RCDT` 동일 기간 | 건수 동일, `SO_VKBUR` 존재·`SO_KUNNR` 부재, 주요 필드값 동일 | N |
| TC-02 | 서비스점조치 + 회수품 접수 | `PA_R02`+`PA_CHX`, `SO_KUNNR` 지정, `SO_RCDT` 동일 기간 | 건수 동일, `SO_KUNNR` 존재·`SO_VKBUR` 부재, RESULTORG='B' 대상 | N |
| TC-03 | 회수품 미접수 | `PA_NCH`, `SO_PRDRE` 동일 기간 | `SO_RCDT` 부재, 접수 완료 건 제외, 건수 동일 | N |
| TC-04 | 유/무상 | `SO_PRTY` = Y / N | Y=유상, N=무상만 조회, 건수·표시값 동일 | N |
| TC-05 | 멀티킷 | `SO_MULTK` = O / X | O→MULTIKITYN='X', X→공백. 뒤로가기 후 재조회 시 조건 유지 | N |
| TC-06 | ALV 화면 구성 | 동일 조회 | SUBS_TYPE 제외 컬럼·순서·컬럼명·Key·편집속성 동일 | N |
| TC-10 | 책임사업장 F4 | RSPC 셀에서 F4 | 동일 목록, 선택 시 RSPC/RSPCT 동일 반영 | N |
| TC-11 | 부적합 사유 Dropdown | NCF_RSN 셀 | 동일 목록, 첫 항목 공백, 코드+설명 결합 | N |
| TC-12 | 품질검사 완료 행 잠금 | ZCLOSE='X' 데이터 | CF_QTY·NCF_QTY·NCF_RSN_VT·RSPC·REMARK 편집 불가 동일 | N |
| TC-14 | 단납 보증기간 | 제조번호 + ZCRM_WRTDG 존재 | 보증기간·회수가능 제조일 동일, CIC 경로 9개월 가산 동일 | N |
| TC-15 | 상세조회 | 1행 선택 → `tbar[1]/btn[23]` | Popup ALV 항목 동일. 0행/2행 선택 시 동일 안내 메시지 | N |
| TC-17 | 변경 없이 종료 | 조회 후 뒤로가기 | 저장확인 Popup 없음, 정상 종료 | N |

### 2차 오류 검증 비교 (저장 차단 — DB 변경 없음)

| ID | 조건 | 기대 결과 |
|---|---|---|
| TC-09-01 | 적합수량+부적합수량 > 요청수량 | 저장 차단, 동일 메시지 |
| TC-09-02 | 부적합수량>0, 사유 공백 | 저장 차단, 동일 메시지 |
| TC-09-03 | 부적합수량=0, 사유 입력(Z10 아님) | 저장 차단 |
| TC-09-04 | 적합수량>0, 책임사업장 공백 | 저장 차단 |
| TC-09-05 | 접수부서/접수자 누락 | 저장 차단 |
| TC-09-06 | 수량 입력, 회수품 입고일자 공백 | 저장 차단 |

### 3차 저장 비교 (DB 변경 있음 — 지정 데이터로만)

| ID | 내용 | 확인 필드 | DB |
|---|---|---|---|
| TC-07 | 적합수량 저장 | CF_QTY, 적합구분, CONFDAT, RSPC, STATUS, RECPTNO, RECPTDATE, NETWR, WAERS, WEARDATS | Y |
| TC-08 | 부적합수량 저장 | NCF='N', NCF_QTY, NCF_RSN_V, 사유텍스트, RECPTNO, RECPTDATE, STATUS, WEARDATS | Y |
| TC-13 | 해지 서비스점 자동 부적합 | AUFSD='Z8' → NCF='N', NCF_RSN_V='Z10' | Y |
| TC-16 | Header 변경 감지 | REMARK1/REMARK2 중 하나만 변경 → GV_MOD_CHECK='X', 저장확인 실행 | Y |

### DB 검증 대상 테이블
- `ZCRM_RECPT`: RECPTNO, RECPTNO_ITEM, RECPTDATE, CF_QTY, NCF_QTY, NCF, NCF_RSN_V, RSPC, STATUS, CONFDAT, NETWR, WAERS, REMARK, REMARK1, REMARK2
- `ZCRM_RECPT_WEAR`: RECPTNO, RECPTNO_ITEM, WEARDATS
- `ZCRM_RECPTNO`: DATUM, SEQNO — 접수번호 채번 형식(일자+순번) 동일 여부, 키 중복 없음

---

## 4. 구독조건 원복 후 테스트

| ID | 내용 | 기대 결과 |
|---|---|---|
| TC-18 | 일반 상담번호 차단 | SUBS_TYPE 공백 상담번호 입력 → 신규에서 미조회 |
| TC-19 | 구독 데이터 없는 상태 실행 | 0건, Dump 없음, ALV 정상 생성, 저장 오동작 없음 |
| TC-20 | SUBS_TYPE='S' 데이터 확인 | 신규에서 조회·변경·검증·저장 정상 |

---

## 5. 실행 순서
1차 비저장(TC-01~06, 10~12, 14, 15, 17) → 2차 오류검증(TC-09-\*) → 3차 저장(TC-07, 08, 13, 16) → 구독조건 원복 → TC-18~20.

## 6. 결과 기록
- 로그: `logs/`, 결과: `results/`. 파일명에 날짜 포함.
- 종료코드: 0 정상 / 1~5 연결 실패 / 99 시스템 가드 / 10 이상 시나리오 실패.
