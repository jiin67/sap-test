# C:\sap-test — SAP GUI Scripting 테스트 규칙

이 폴더는 ABAP 프로그램(개선분) 을 SAP GUI Scripting(VBScript) 으로 테스트하는 작업 공간이다.

## 1. 환경 (2026-09-08 확인)

| 항목 | 값 |
|---|---|
| SAP GUI | 7.70 (Scripting Engine 7700.257.0), `C:\Program Files (x86)\SAP\FrontEnd\SAPgui` |
| 실행 방법 | `cscript //nologo <script>.vbs` (PowerShell) |
| 연결 확인 스크립트 | `sap_scripting_check_kr.vbs` — 종료코드 0 이면 정상 |
| 등록된 시스템 | KDD(개발) 10.100.8.116 / KDQ(QA) 10.100.8.84 / KDP(운영) 10.100.8.7 |
| 현재 세션 | KDQ 클라이언트 210, 사용자 KKR003-1, 언어 KO, `/app/con[0]/ses[0]` |

## 2. 절대 규칙 (안전)

- **KDP(운영)에는 어떤 스크립트도 실행하지 않는다.** 모든 스크립트는 시작 시 `session.Info.SystemName` 을 읽어 `KDQ` 또는 `KDD` 가 아니면 즉시 `WScript.Quit` 한다.
- 로그온/암호 입력은 스크립트로 하지 않는다. 사용자가 SAP Logon 으로 직접 로그온한 **기존 세션에 attach** 만 한다.
- DB 를 변경하는 트랜잭션(저장/전기/릴리스 등)은 시나리오 문서에 명시된 경우에만, 지정된 테스트 데이터로만 실행한다.
- 테스트 데이터는 시나리오에 정의된 식별자(예: 텍스트 필드에 `ZTEST_YYYYMMDD` 접두어)를 붙여 사후 식별·정리가 가능하게 한다.
- 스크립트 실행 전 반드시 `sap_scripting_check_kr.vbs` 로 연결을 확인한다.
- 실행 중인 세션의 트랜잭션을 가로채지 않는다. 시작은 항상 `/n<TCODE>` 로 화면을 초기화하고, 종료 시 `/n` 으로 초기화면에 돌려놓는다.

## 3. VBScript 작성 규칙 (여기서 실제로 겪은 문제 기준)

- **파일 인코딩은 CP949(ANSI)** 로 저장한다. UTF-8(BOM 없음)로 저장하면 cscript 가 한글 문자열을 깨뜨려 "종료되지 않은 문자열 상수" 오류가 난다.
  - PowerShell 저장 예: `[IO.File]::WriteAllText($p,$t,[Text.Encoding]::GetEncoding(949))`
- **컬렉션 인덱스는 `Children.ElementAt(i)` 를 쓴다.** `Children(i)` 는 For 루프의 정수형 변수를 넘기면 오류 424(객체가 필요합니다)가 난다.
- 세션 attach 표준 코드:
  ```vbs
  Set SapGuiAuto = GetObject("SAPGUI")
  Set application = SapGuiAuto.GetScriptingEngine
  Set connection = application.Children.ElementAt(0)
  Set session = connection.Children.ElementAt(0)
  If session.Info.SystemName <> "KDQ" And session.Info.SystemName <> "KDD" Then WScript.Quit 99
  ```
- `On Error Resume Next` 는 구간별로만 쓰고, 각 findById 뒤에 `Err.Number` 를 확인한다. 오류 무시 상태로 다음 화면 조작을 계속하지 않는다.
- 화면 전환 후에는 `session.findById("wnd[0]/sbar").Text` (상태바 메시지) 와 `session.Info.ScreenNumber` 를 읽어 기대 화면인지 검증한 뒤 다음 단계로 간다.
- 팝업(`wnd[1]`) 은 존재 여부를 먼저 확인하고 처리한다.
- 결과 판정은 화면 제목 문자열이 아니라 **필드 값·상태바 메시지·ALV 셀 값** 으로 한다. 제목은 로그온 언어(KO)에 따라 바뀐다.

## 4. 실행·로그 규칙

- 콘솔 한글이 깨지므로 출력은 파일로 받는다:
  ```powershell
  [Console]::OutputEncoding=[Text.Encoding]::GetEncoding(949)
  cscript //nologo scripts\xxx.vbs 2>&1 | Out-File -Encoding utf8 logs\xxx_yyyyMMdd_HHmm.txt
  ```
- 로그는 `logs/`, 테스트 결과(예상 vs 실제)는 `results/` 에 남긴다. 파일명에 날짜를 넣는다.
- 종료코드 규약: 0 정상 / 1~5 연결 단계 실패 / 99 시스템 가드 차단 / 10 이상 시나리오 실패.

## 5. 폴더 구조

```
C:\sap-test
├── CLAUDE.md                     이 파일
├── sap_scripting_check_kr.vbs    연결 확인
├── scripts/                      시나리오별 VBS (S01_xxx.vbs ...)
├── scenarios/                    시나리오 문서 (md), 템플릿: _TEMPLATE.md
├── logs/                         실행 로그
└── results/                      결과 보고
```

## 6. 시나리오 작성 절차

1. `scenarios/_TEMPLATE.md` 를 복사해 프로그램별 시나리오 문서를 만든다.
2. 개선 전/후 동작 차이를 먼저 적고, 그 차이를 검증하는 케이스(정상·경계·오류)를 나열한다.
3. 케이스마다 입력값, 실행 경로(T-code, 화면 순서), 기대 결과(필드/메시지/ALV 값)를 확정한 뒤 VBS 를 만든다.
4. VBS 는 케이스 1개 = 서브루틴 1개로 만들고, 결과를 `PASS/FAIL` 로 로그에 남긴다.

## 7. 세션 운용 패턴 (검사 스크립트에서 검증됨)

- 사용자 화면을 건드리지 않으려면 `base.createSession` 으로 **새 세션을 열어** 작업하고, 끝나면 `session.findById("wnd[0]").close` 로 그 세션만 닫는다. `/nex` 는 접속 전체를 끊으므로 절대 쓰지 않는다.
- 스크립트가 중간에 죽으면 세션이 남는다. 시작 시 `connection.Children.Count > 1` 이면 마지막 세션을 재사용한다.
- ALV 그리드는 `Type="GuiShell"`, `SubType="GridView"` 다. 타입으로 찾을 때 SubType 을 봐야 한다.
- 도킹 컨테이너 ALV 는 `wnd[0]/usr` 가 아니라 `wnd[0]/shellcont/...` 아래에 있다. usr 가 비어 보여도 wnd[0] 전체를 봐야 한다.
- F4 값 목록 팝업은 `wnd[1]/usr/lbl[x,y]` 라벨로 읽힌다.
- ALV 툴바 버튼의 Text/Tooltip 은 한글이 깨져 나온다. 버튼은 `GetToolbarButtonId(i)` 로 얻은 기능코드를 `grid.pressToolbarButton "CODE"` 로 누른다.
- 로그 파일은 `FileSystemObject.CreateTextFile(path, True, True)` 로 UTF-16 으로 쓰고, 읽을 때 `iconv -f UTF-16LE -t UTF-8` 또는 `Get-Content -Encoding Unicode` 를 쓴다.

## 8. ZSD600 (Z33SDR600) 화면 참조 — 2026-09-08 KDQ 210 기준

제목: `[RU] 1C I/F Sales Order Management` / 개발클래스 Z33SD / KDQ 활성버전: TR KDDK999712 (INFO10, 2026-06-18 12:31)

### 선택화면 (screen 1000) 필드 ID — 모두 `wnd[0]/usr/` 하위
| 항목 | ID | 비고 |
|---|---|---|
| 수행옵션 | `radP_RBLIST` 주문조회 / `radP_RBLOG` 주문로그 / `radP_RBSYNC` 데이터 동기화 | 라디오 바꾼 뒤 Enter(sendVKey 0) 해야 화면이 갱신됨 |
| DB 연결명 | `ctxtP_DBCON` | 동기화 모드에서만 표시. 기본값 `CMSRUS_QA` |
| 1C Unique ID | `txtS_ID-LOW` / `txtS_ID-HIGH` | |
| 외부 주문번호 | `txtS_BSTKD-LOW` / `-HIGH` | |
| 버전 | `txtS_VERSN-LOW` / `-HIGH` | 조회·로그 모드만 |
| 생성일(1C) | `ctxtS_ERDAT-LOW` / `-HIGH` | 기본값 어제~오늘 |
| 변경일(1C) | `ctxtS_AEDAT-LOW` / `-HIGH` | |
| 삭제된 항목 조회 | `chkP_LOEVM` | 조회·로그 모드만 |
| 주문상태 | 조회 `ctxtS_STAT1-*` (기본 A~B) / 로그 `ctxtS_STAT2-*` (빈값) / 동기화 `ctxtS_STAT3-*` (기본 C, 옵션버튼 있음) | 모드마다 다른 필드 |
| 자재번호 | `ctxtS_MATNR-LOW` / `-HIGH` | |
| 실행 | `wnd[0]/tbar[1]/btn[8]` (F8) | |

주문상태 도메인: **A = 1C 입력, B = SAP 동기화, C = 주문 마감**

### 결과화면 (screen 100) — 도킹 스플리터, 상단 헤더 그리드 + 하단 품목 그리드
- 헤더 그리드: `wnd[0]/shellcont/shell/shellcont[0]/shell`
  - 조회 모드 31컬럼: STATU, ID, BSTKD, VERSION, VBELN, VBELN_VL, AUART, VKORG, VTWEG, SPART, KUNNR, VKBUR, VKGRP, VDATU, REMARK, MONTH, NAME1, NETPR_SUM, CREDAT, CREZET, CREZON, CHGDAT, CHGZET, CHGZON, ERNAM, ERDAT, ERZET, AENAM, AEDAT, AEZET, LOEVM
  - 로그 모드 26컬럼: ID, BSTKD, VERSION, AUART … STATU, ERNAM, ERDAT, ERZET, AENAM, AEDAT …
  - 조회 모드 사용자 버튼(기능코드): `CREA_ORD` 판매오더 생성, `ELIM_ORD` 판매오더 삭제, `CLOS_ORD` 주문 마감, `CLOS_CANC` 마감 취소, `CRE_DELIVERY` 납품 생성/삭제(같은 코드 2개, 순서 18·19), `DELE_SYNC` 동기화 데이터 삭제
- 품목 그리드: `wnd[0]/shellcont/shell/shellcont[1]/shell`, 12컬럼: ID, BSTKD, SEQNO, MATNR, MAKTX, ZMENG, MEINS, NETPR, WAERS, KDKG1, LGORT, SER_BTN. 사용자 버튼 `CHANGE_MAT`. 처음엔 0행(헤더 선택 시 채워지는 것으로 추정)
- **DB 를 바꾸는 버튼**: CREA_ORD, ELIM_ORD, CLOS_ORD, CLOS_CANC, CRE_DELIVERY, DELE_SYNC, CHANGE_MAT, 그리고 동기화 모드 실행(F8). 시나리오에 명시된 데이터로만 누른다.

### 데이터 현황 (2026-09-08)
- 생성일(1C) 2026.01.01~09.08, 상태 전체: 조회 2,208건 / 로그 2,232건. 데이터는 러시아어(키릴) 포함, VKORG 6100, AUART YOR.
- 기본 날짜 범위(어제~오늘)로는 0건 → 케이스 입력값에 날짜를 반드시 넣는다.
- P_DBCON F4 적중 리스트(AIRWMS, KWMSDB, KWMSDB_DEV, NPLM, PLMQA)에 `CMSRUS_QA` 가 없다. 동기화 테스트 전에 DBCO 에 연결이 있는지 확인 필요.

## 9. ZCRM052 (ZCRMR052) / ZCRMR052_S 화면 참조 — 2026-09-09 KDQ 210 기준

제목: `회수품적합성관리`. 기존 T-code `ZCRM052` → 프로그램 `ZCRMR052`.
구독 전용 신규 프로그램 `ZCRMR052_S` 에는 임시 전용 T-code 가 있다.
**실제 T-code 값은 `scripts/config_local.txt` 의 `TCODE_NEW` 에만 두고, 문서·보고서에는 `ZCRM052_S` 로 표기한다.**
스크립트는 `NewTcode()` 로 그 값을 읽어 `/n<tcode>` 로 기동하고, 없으면 SE38 실행으로 대체한다.
SE38 로 기동하면 tbar[1] 에 ABAP 편집기 버튼(디버깅·이름바꾸기 등)이 섞이므로 비교 테스트에는 T-code 를 쓴다.

### 선택화면 (screen 1000) — `wnd[0]/usr/` 하위, 두 프로그램 동일
| 항목 | ID | 비고 |
|---|---|---|
| 처리조직 | `radPA_R01` 센터조치 / `radPA_R02` 서비스점조치 | 변경 후 Enter 필요 |
| 센터 | `ctxtSO_VKBUR-LOW` / `-HIGH` | PA_R01 일 때만 **존재** |
| 서비스점 | `ctxtSO_KUNNR-LOW` / `-HIGH` | PA_R02 일 때만 **존재** |
| 회수품요청일자 | `ctxtSO_PRDRE-LOW` / `-HIGH` | 항상 존재 |
| 상담번호 | `txtSO_CSLTN-LOW` / `-HIGH` | |
| 접수일자 | `ctxtSO_RCDT-LOW` / `-HIGH` | PA_CHX 일 때만 존재, 기본값 오늘~오늘 |
| 제품 계층구조 | `ctxtSO_PRDH-LOW` / `-HIGH` | |
| 유/무상 | `ctxtSO_PRTY-LOW` | 단일값 |
| 자재번호(부품) | `txtSO_PRT-LOW` / `-HIGH` | |
| 멀티킷 사용 여부 | `ctxtSO_MULTK-LOW` | 단일값 |
| 접수여부 | `radPA_CHX` 회수품 접수 / `radPA_NCH` 회수품 미접수 | |

> 조회조건은 비활성(disabled)이 아니라 **화면에서 사라진다.** 활성/비활성 판정은
> `Changeable` 이 아니라 `findById` 성공 여부(존재 여부)로 해야 한다.

### 결과화면 (screen 100)
- 접수부서 `wnd[0]/usr/txtGS_HEADER-REMARK1`, 접수자 `wnd[0]/usr/txtGS_HEADER-REMARK2`
- ALV `wnd[0]/usr/cntlCON1/shellcont/shell` (GuiShell / GridView) — ZSD600 과 달리 도킹 아님
- 상세조회 버튼 `wnd[0]/tbar[1]/btn[23]`
- **조회 건수는 창 제목 `회수품적합성관리 (n)` 의 괄호 안 숫자로 읽는다.**

### 데이터 규모 (2026-09-09)
- 센터조치 + 회수품 접수, 접수일자 2026.01.01~09.09 → **71,912건**. 하루 평균 약 290건.
- 광범위 조회는 수 분 이상 걸린다. 비교 테스트는 **1~2일 단위**로 잡는다.

## 10. 스크립팅 엔진 분리 사고 (2026-09-09)

장시간 조회 중인 `cscript` 를 `Stop-Process` 로 강제 종료했더니, 이후
`GetObject("SAPGUI")` 가 오류 432 로 실패했다. cscript 와 PowerShell 양쪽 동일.
SAP GUI 프로세스 자체는 정상(응답함, 모달 없음)이고 세션도 살아 있었다.

- **장시간 실행 중인 스크립트를 강제 종료하지 않는다.** 끝날 때까지 기다린다.
- 조회 스크립트는 처음부터 **좁은 조건**으로 만들어 장시간 실행을 피한다.
- 이미 발생했다면: SAP 창 전체 종료 → SAP Logon 재시작 → 사용자가 직접 재로그온
  → `sap_scripting_check_kr.vbs` 로 확인. 스크립트로 로그온하지 않는다.

## 11. SE38 관련 실측

- 속성 화면 라디오는 `radRS38M-FUNC_HEAD` (`FUNC_ATTR` 아님). 조회 버튼은 `btnSHOP`.
- **ABAP 소스 본문은 GUI 스크립팅으로 읽을 수 없다.** 편집기 shell 의 `Text` 는
  19자만 반환한다. 소스 확인이 필요한 검증은 개발자에게 요청한다.

## 12. SE16N 사용 규칙 (2026-09-09 실측)

화면 ID:

| 항목 | ID |
|---|---|
| 테이블 | `wnd[0]/usr/ctxtGD-TAB` |
| 최대 적중 수 | `wnd[0]/usr/txtGD-MAX_LINES` (기본 500) |
| **엔트리 유지보수(편집)** | `wnd[0]/usr/chkGD-EDIT` |
| 조건 테이블 | `wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC` |
| 조건 필드명(설명) | `.../txtGS_SELFIELDS-SCRTEXT_M[0,r]` |
| 조건 값 From / To | `.../ctxtGS_SELFIELDS-LOW[2,r]` / `-HIGH[3,r]` |
| 실행 | `wnd[0]/tbar[1]/btn[8]` (F8) |
| 결과 ALV | `wnd[0]/usr/cntlRESULT_LIST/shellcont/shell` |

- **`chkGD-EDIT` 가 켜진 채로 남아 있을 수 있다.** 2026-09-09 첫 실행에서 실제로 ON 이었다.
  SE16N 을 쓰는 스크립트는 **시작할 때마다 이 체크박스를 확인하고 반드시 끈다.**
  켜진 상태로 실행하면 결과 화면에서 테이블을 직접 편집할 수 있게 되어 위험하다.
- 조건 테이블에는 기술적 필드명이 안 보이고 **설명(SCRTEXT_M)만** 보인다.
  조건은 설명 문자열로 찾아 넣는다. 행이 많으면 `VerticalScrollbar.Position` 으로 스크롤한다.
- 결과 ALV 의 컬럼명은 기술적 필드명이라 그대로 쓰면 된다.

## 13. ALV 대량 행 읽기

`GetCellValue(row, col)` 는 **화면에 로드된 행만** 읽힌다. 500행 이상을 덤프하려면
`VisibleRowCount` 만큼 끊어서 `FirstVisibleRow` 를 옮겨가며 읽어야 한다.
안 그러면 화면 밖 행이 빈 값이나 오류로 나온다.

## 14. 화면 캡처 (테스트 증거)

`session.findById("wnd[0]").HardCopy path, 2` 로 PNG 를 저장한다. 1929x1368, 약 130KB.
ImageType: 0=BMP(8MB), 1=JPG, 2=PNG, 3=GIF. **PNG(2) 를 쓴다.**
팝업은 `wnd[1]).HardCopy` 로 따로 찍는다.
별도 세션에서 찍히므로 사용자 화면을 건드리지 않는다. 결과는 `results/shots/` 에 둔다.

## 15. FBGENDAT / SU3 / 소스스캔 화면 참조 — 2026-09-12 KDD 900 기준

- KDD 접속은 `application.Children` 을 돌며 `Info.SystemName = "KDD"` 인 접속을 고른다(순서 보장 없음). KDD 는 클라이언트 900, 사용자 INFO10.
- **FBGENDAT** (SE38 실행, 제목 `구성: 테스트데이타생성`): `txtPA_FNAME` 기능모듈, `txtPA_MODE` 모드(A/B/C/D, RFC 는 **B**),
  라디오 `radPA_INS` 해당엔트리활성화 / `radPA_READ` 활성엔트리조회 / `radPA_DEL1` 해당엔트리삭제 / `radPA_DELA` 모든엔트리삭제. 실행 F8.
  결과 ALV `wnd[0]/usr/cntlGRID1/shellcont/shell` 컬럼 FNAME, MODE. 등록 시 CTS 팝업 없음(이관 안 됨). 기능모듈 필드는 마지막 값을 기억한다.
- **SU3** 매개변수 탭: `tabsTABSTRIP1/tabpPARAM`, 그리드 `.../ssubMAINAREA:SAPLSUID_MAINTENANCE:1104/cntlG_PARAMETER_CONTAINER/shellcont/shell`
  컬럼 PARID, PARVA, PARTXT. 빈 행에 `ModifyCell` 후 Enter → 저장 `tbar[0]/btn[11]`. 저장하면 SAP Easy Access 로 빠져나온다.
- **RS_ABAP_SOURCE_SCAN**: `txtSSTRING-LOW` 문자열, `ctxtFUNCGRP-LOW` 기능그룹(와일드카드 가능), `chkINCLU` Explode INCLUDEs.
  결과는 리스트(SAPMSSY0/120) 라벨 `lbl[4,y]` 오브젝트명·라인, `lbl[11,y]` 소스. GUI 로 못 읽는 소스 확인은 이걸로 한다.
- SE37 테스트 초기화면 `테스트데이타 디렉토리` 버튼 = `wnd[0]/tbar[1]/btn[42]`. 없으면 상태바 `테스트 데이타 디렉토리가 없습니다`.
- FBGENMAC 사내 적용 패턴: 기능그룹 TOP 에 `INCLUDE fbgenmac.`, 각 FM 선언부 앞에 `fbgenmac 'FM명'.` (ZCRMG02, Z33EDIG01 참고).

## 16. ZCRMR194 / ZCRM052 계열 참조 — 2026-09-14 KDQ 210 기준

- **ZCRMR194** 는 T-code 없음 → SE38 실행. 선택화면 `ctxtS_VKORG-LOW` / `ctxtS_VKBUR-LOW` / `txtS_CSLTNO-LOW,HIGH` / `ctxtS_INDAT-LOW`. 결과는 리스트(SAPMSSY0/120).
  ISFREE='Y' 건은 청구(ZCRM_REQ) 뒤에 **무상오더 ZFD1 → VL01N 납품 → VL02N 출고전기까지 실제로 만든다**(재고 이동). 가용재고 없으면 ZCRM_AUTO_ERRLOG 만 남음. 반드시 S_CSLTNO 로 대상 제한.
- T-code: ZCRM052·ZCRM052S·ZCRM053 → ZCRMR052 (sy-tcode 로 구독/조회전용 분기). ZCRM053S 는 2026-09-14 에 TSTC 가 ZCRMR053(없는 프로그램)이라 덤프였고 **2026-09-18 에 ZCRMR052 로 수정 확인**.
- ZCRM052 계열 결과 ALV 레이아웃 키가 T-code 공통 → 09-14 에는 기본 레이아웃이 구독구분을 숨겼고, 09-18 재테스트에서는 반대로 **비구독 ZCRM052·053 에도 구독구분 컬럼이 표시**된다(레이아웃이 NO_OUT 을 덮음, 미해결). 컬럼 존재 확인은 `&MB_VARIANT`→`&COL0` 팝업의 숨긴 열 그리드로.
- 결과 ALV 증적 캡처는 가로 스크롤(FirstVisibleColumn/setCurrentCell) 후 HardCopy 가 행을 빠뜨린다. **행 선택 후 `&DETAIL`(세부사항) 팝업을 캡처**한다.
- SE16N 편집으로 INT 필드에 **음수(-3, 3-)를 넣으면 오류 없이 저장이 안 된다**(변경 완료 팝업 없음).
- 스크립트: `S19_se16n_insert.vbs`(허용 상담번호만 새 행 삽입, 원본은 조회 TSV 로만 읽음), `S20_flow.vbs`(범용 스텝 실행기: TC/SET/SELECT/BTN/TBTN/CELL/DUMP/STATUS/GRIDMENU/POPGRIDS/POPSCROLL/SHOT…),
  `S12_se16n_cols.vbs` 는 환경변수 `S12_SHOT=1` 이면 결과 화면 캡처, `S18_source_scan.vbs` 는 `chkP_REGEX=X` `txtPLUSMINU=0` 로 정규식 "." 을 주면 전체 소스 덤프(`parse_scan.py` 로 정리).
## 17. ZSD600 재테스트에서 확정된 동작 (2026-09-15)

- **자재변경(CHANGE_MAT) '저장'(wnd[1]/tbar[0]/btn[8])은 메모리 반영**이다. Z33SDT641/Z33SDT644 는 바뀌지 않고 `/n` 으로 나가면 사라진다. 이력(Z33SDT644)은 CREA_ORD 시점에 기록. 검증은 **재조회(F8) 없이** 같은 화면의 품목 ALV 를 읽는다.
- 자재변경 대화상자 검증은 **Enter 시 MARA 만**(오류는 `wnd[2]` 팝업 "오류", 본문 `wnd[2]/usr/txtMESSTXT1`). 동일 자재 재입력은 메시지 없이 닫힌다.
- 납품 생성/삭제·판매오더 생성/삭제 버튼은 **예/아니오 확인창**(SAPLSPO1, `wnd[1]/usr/btnBUTTON_1`=예, `btnBUTTON_2`=아니오). F12 는 비활성이라 **아니오 버튼을 눌러** 취소한다.
- 실행 결과는 상태바 "요청을 완료했습니다. 생성로그를 확인하십시오" + **결과 로그 팝업** `wnd[1]/usr/cntlGRID1/shellcont/shell` (컬럼 BSTKD, VBELN_VA, VBELN_VL, RETCODE, RESULT, RESULT_MESSAGE …).
- 단일 주문만 다룰 때는 선택화면 **`txtS_ID-LOW`(1C Unique ID)** 로 조회해 헤더 1행만 띄운다(BSTKD 는 버전별로 여러 행).
- 헤더 ALV 도 §13 그대로: 16행 이상이면 `FirstVisibleRow` 로 넘기며 읽어야 한다. 안 그러면 뒤 행이 빈 값으로 나와 "중복"처럼 보인다.
- SE16N 조건 입력: 설명(SCRTEXT_M) 매칭 — ID 필드는 `Uniq`, 자재번호는 `자재번호`. `chkGD-EDIT` 는 매번 끈다.
- **미해결**: PNGB3500035L001 은 MARC 6100/02 에 있는데도 ZSD600 ATP(ZCRMF02_RFC_MAT_AVAILABILITY)가 "마스터 레코드 없음"으로 실패 → CREA_ORD 불가. 판매오더 필요 케이스는 이 원인 해결 후 재개.
- 이 폴더에서 새 VBS 는 **순수 ASCII + ChrW()** 로 쓴다(한글 리터럴 인코딩 문제 회피). 한글은 SAP 에서 읽어온 값만 로그에 남긴다.

## 18. ZCRM171 / ZCRM171S (ZCRMR171) 참조 — 2026-09-18 KDQ 210 기준

- 선택화면: `radPA_TAB` 조회(기본, 요청일자 `ctxtSO_CADAT-LOW/HIGH` 필수) / `radPA_NEW` 정산(`txtPA_SPMON` 기간월 = ZCRM_RECPT-CONFDAT 의 월, `ctxtSO_VKBUR-LOW`). 결과 ALV `wnd[0]/shellcont/shell`.
- 정산 화면 버튼: `tbar[1]/btn[2]` 본사확정, `btn[5]` 본사확정취소, `btn[46]` 전표생성(확인 → F4_DATE 달력 `wnd[1]`, 오늘=`tbar[0]/btn[0]`), `btn[45]` 전표취소. 상태 ZICON `@5C@` 빨강 / `@5D@` 노랑 / `@5B@` 초록.
- 달력에서 다른 날짜: `wnd[1]/usr/cntlCONTAINER/shellcont/shell` 의 `FocusDate`·`SelectionInterval` 설정(설정 즉시 팝업이 닫히고 진행됨, 읽기는 445 오류). S20 `CALDATE:yyyymmdd`.
- 대상: ZCRM_RECPT ⨝ ZCRM_CICSVC(PRICETYPE N, RESULTORG B) ⨝ MVKE(ASKR/10). 테스트 데이터는 CICSVC+REQ 를 SE16N 삽입 후 ZCRM052/052S 에서 접수 저장하면 만들어진다. ZCRM_RECPTF 키 = SPMON+KUNNR_SVC+SUBS_TYPE, **SE16N 편집 불가**, NETWR 는 INT4(43,450 → 435).
- **2026-09-18 22시 재테스트: 아래 결함 모두 수정 확인**(재검증 실패 시 RETURN, KUNCR/KUNCC 2단계 LOOP, 정산 ALV 에 Message(ZMSG) 컬럼 추가 → 컬럼 순서가 SUBS_TYPE, ZICON, … 으로 바뀜). 결과서 v0.02.
- (이력) 2026-09-18 새벽 확인된 결함: SET_FIPOS/SET_FIDEL 재검증 실패 시 `EXIT` 가 LOOP 만 빠져나가 전표 생성·취소가 그대로 진행. SET_KUNCR/SET_KUNCC 다중 선택 시 부분 저장·삭제, 재검증 메시지가 S007 에 덮임.
- 스크립트: `S21_se16n_setfield.vbs`(테스트 상담건 SUBS_TYPE 토글, 1행·이전값 검증), S20 에 `EXEC:`(조회 화면을 유지한 채 다른 스크립트 실행)·`SELROWS:col=v1|v2`·`CELLALL:`·`CALDATE:` 추가, TC 스텝에서 창 최대화. 래퍼 `t171.ps1`, DB 확인 `db171.ps1`, FI 확인 `fi171.ps1`(BSEG 조회는 수 분 걸릴 수 있음).
- 해지 대리점(예 100035·100036·100038·100039)은 ZCRM052 에서 부적합사유 Z10 이 자동으로 붙는다. 테스트 대리점은 100032·100034·100037·100090·100273 처럼 정상점을 쓴다.- **2026-09-18 23시 2차 데이터 전체 재수행(v0.03)**: 대리점 100041·043·044·045·047·050(정상점, KNVV-AUFSD 공백; `Z8` 이면 해지점), 상담번호 2026091892001~008. 기존 49건 PASS 유지. **신규 결함 X-04**: 조회(상세) ALV 와 정산 ALV 가 레이아웃 키를 공유해, 정산 쪽에서 저장된 기본 레이아웃이 상세조회 컬럼 21개를 숨김(5개만 표시). 통합 결과서 `…_통합_v0.03_20260918.xlsx`(개정이력 0.01~0.03, 생성 스크립트 make_fill_data_171.py → 171b → 171m → 171n).
## 19. ZCRMR177_1 (ZCRM177A / ZCRM177AS) 참조 — 2026-09-19 KDQ 210 기준

- T-code: ZCRM177A(기존) / **ZCRM177AS(신규, TSTC 등록 확인)** → 둘 다 프로그램 ZCRMR177_1 / 화면 1000. 결과화면 제목은 `AS 비용처리` vs `[구독] AS 비용처리`(SET TITLEBAR 'ZSUB').
- 선택화면: `ctxtP_BUDAT`(전기일자) `ctxtP_BUKRS` `ctxtP_VKORG`(ASKR) `ctxtSO_VKBUR-LOW/HIGH`(센터) `txtSO_CSLTN-LOW/HIGH`(상담번호) `txtP_TECHNM`(기사명) **`ctxtSO_INDAT-LOW/HIGH`(조치일자 = ZCRM_CICCOST_M-INDATE)** `chkP_COMP`(완료건 포함). 결과 ALV `wnd[0]/shellcont/shell/shellcont/shell`.
- 데이터 원천: `ZCRM_CICCOST_M a` ⨝ `ZCRM_CICSVC c`(CLMART='A') ⨝ LEFT `ZCRM_CICCOST_PAY b`. **Native SQL 없음.** `ZCRM_CICCOST_SUB` 에 COST=0 행이 없으면 상담번호가 통째로 빠진다(F01 165 "올유상인건 지운다") → 테스트 데이터는 SVC + COST_M + SUB 3종 세트가 필요.
- 구독 분기: `GR_SUBS` 레인지(구독 I/EQ 'S', 일반 E/EQ 'S') + `c~subs_type IN gr_subs`. 표시는 `SUBS_TEXT`(유상/무상/구독무상(N년이후)/구독무상(N년이내)), 분기 기준 `SUBSWRTMONTH > 0` 아니면 이내.
- `ZCRM_CICSVC-SUBSWRTMONTH` 는 **INT2(부호 있음)** 이지만 SE16N 편집으로 음수(-6, 6-)를 저장할 수 없고 SM30 유지보수 화면도 없다 → QA 에서 음수 데이터 생성 불가. `SUBS_TYPE` 도메인은 공백/S 만 허용('X' 저장 거부).
- CIC 적재는 **`ZCRMF02_RFC_ZCRM_CICSVC_ALL`(기능그룹 ZCRMG02)** 가 담당한다(화면의 `상담번호 동기화(CIC->SAP)` 버튼 = ucomm SYNC → F01 3065). 시나리오의 Native SQL·연동 항목은 전부 이 FM 소관이고 ZCRMR177_1 에는 EXEC SQL 이 없다. 소스 정리본 `results\src\ZCRMG02_subs_blocks.txt`.
  - `LZCRMG02U27`(ZCRM_CICSVC) 는 `:GT_ZCRM_CICSVC-SUBSWRTMONTH` 직접 수신, `LZCRMG02U73`(ZCRM_CICSVC_NEW) 는 이름이 32자로 30자 제한을 넘어 보조 변수 `:GV_SUBSWRTMONTH`(LZCRMG02TOP 35, `LIKE ZCRM_CICSVC_NEW-SUBSWRTMONTH` → INT2) 로 받은 뒤 F01 237 에서 대입. **부호는 전 구간 보존된다.**
  - 요구사항의 WRTDATE 도 두 경로 모두 적재된다(U27 281↔356, U73 252↔314). 화면 프로그램이 WRTDATE 를 읽지 않는 것은 정상(구독여부 판정은 SUBSWRTMONTH 만 사용).
- **결함(미해결)**: (D-01) 전역 기본 레이아웃 `/PT` 가 신규 '구독여부' 컬럼을 숨김 — `GV_VARIANT-REPORT = SY-REPID` 만 써서 두 T-code 가 레이아웃 키 공유(ZCRM052·ZCRM171 의 B-03/X-04 와 같은 유형). (D-02) 다품목 상담번호에 구독/비구독이 섞이면 `DELETE ADJACENT DUPLICATES ... COMPARING csltno` 때문에 양쪽 화면에 중복 표시되고 구독여부가 화면마다 달라짐.
- **ALV 선택 체크박스(CHK)는 GUI 스크립팅으로 프로그램에 전달되지 않는다.** modifyCell/modifyCheckbox + Enter/열이동/행선택 조합 모두 `LOOP AT gt_data WHERE chk='X'` 에 걸리지 않아 항상 "선택한 데이터가 없습니다" → 문서처리(주문/납품/출고) 계열은 수동 테스트 필요.
- **`GuiFrameWindow.maximize` 를 호출한 세션에서는 이후 `HardCopy` 가 오류 -2147418113 으로 실패한다**(파일 미생성). S20 의 TC 스텝 최대화는 환경변수 `S20_NOMAX=1` 로 끌 수 있다. 캡처가 안 될 때 이걸 먼저 확인한다.
- 스크립트: `S23_se11_field.vbs`(SE11 필드 자료형 조회), S20 에 `POPSELCOL:`(팝업 그리드에서 열 이름으로 행 선택)·`CHKBOX:`·`SELCOL:` 추가. 래퍼 `t177.ps1`(-Mode A|S, -From/-To, -Cslt, -Comp, -Subs). 데이터 생성 `mk177data.ps1 -Part svc|m|sub|doc`.
- **2026-09-19 오후 D-01·D-02 수정 재테스트(v0.03)**: 둘 다 해결. D-01 은 기본 레이아웃 `/PT` 에 구독여부를 표시열 2번째로 재저장(코드 변경 없음 → 레이아웃 키 공유는 잔존). D-02 는 개선안 3 적용 - F01 116 SQL 필터 주석, 150 `SORT gt_data BY csltno ASCENDING subs_type DESCENDING subswrtmonth ASCENDING`, 152~161 혼재 판정 LOOP, 164 중복 제거, 167~171 모드별 DELETE. 혼재 상담건은 구독 화면에만, 합계 3+7=10.
  - **신규 D-04**: `TEXT-M90` 텍스트 기호가 없어(SE38 텍스트요소: M01~M12, M16~M19 뒤 바로 T04) 혼재 메시지가 공백. 텍스트 기호 생성 필요.
  - **신규 D-05**: 혼재 판정 LOOP 가 행마다 `SELECT COUNT( DISTINCT )` 실행(중복 제거 전이라 품목 단위). ZCRM177A 73→87초, ZCRM177AS 19→29초. GROUP BY 1회 조회로 대체 권고.
  - 스크립트 추가: `S24_se38_textsym.vbs`(SE38 텍스트기호 목록 조회·검색).
- 결과서: `results\SAP_KR_단위테스트케이스및결과서_ZCRMR177_1_구독조회분리_v0.03_20260919.xlsx` (51건: PASS 29 / FAIL 4 / 미실행 18, 개정이력 0.01~0.03)

## 20. SR03000035400 전체 프로세스 (2026-09-19 E2E 통합 테스트)

- **개선 대상 오브젝트 전수 확인**(RS_ABAP_SOURCE_SCAN, `Z*` 리포트 + `Z*` 기능그룹에서 "SR03000035400" 검색):
  ZCRMR052(+C01/F01/O01/T01), ZCRMR177_1(+includes), ZCRMR171(+ZCRMR0171_1*/_2F01), ZCRMR058_1, ZCRMR060_1, 기능그룹 ZCRMG02(LZCRMG02 TOP/F01/U27/U73/U94), ZCRMG03(TOP - `INCLUDE fbgenmac.` 한 줄).
  **ZCRMR194 는 대상 아님**(표식 0건, 테스트 데이터 생성용으로만 사용).
- 프로세스: CIC → ①ZCRMG02 적재(ZCRM_CICSVC / ZCRM_CICSVC_NEW) → ②ZCRM177A/AS 센터 무상 매출(오더·납품·출고) / ③ZCRM052·052S 회수품 접수(ZCRM_RECPT) → ④ZCRM171·171S 정산·본사확정·FI(ZCRM_RECPTF, 키+SUBS_TYPE) / ⑤ZCRM058A·S 무상지원금(ZCRM_ASSPTFEEH1, 키+SUBS_TYPE) → ⑥ZCRM060A·S 상세·전송.
- 공통 패턴: `sy-tcode` 로 모드 판별 → RANGES(구독 I/EQ 'S', 일반 E/EQ 'S') → 결과 테이블 키에 SUBS_TYPE 추가 → 화면은 내부코드 대신 업무 문구.
- **ZCRM052 접수 저장 절차**: 서비스점조치(radPA_R02) + 회수품 미접수(radPA_NCH) + 상담번호 → 실행 → ALV 편집열 `CF_QTY`(적합수량) `RSPC`(책임사업장 1100) **`WEARDATS`(회수품입고일자, 빠뜨리면 E 오류)** + 헤더 `txtGS_HEADER-REMARK1/2` → `tbar[0]/btn[11]` → 확인 팝업 `wnd[1]/usr/btnBUTTON_1`. ZCRM_RECPT 의 SUBS_TYPE·KUNNR_SVC 는 채워지지 않는다(기존 동작, 정산은 ZCRM_CICSVC 에서 다시 읽음).
- ZCRM171 전표생성은 본사확정 직후 선택이 풀리므로 **행을 다시 선택**해야 한다.
- 058/060 선택화면: `radPA_NEW`(전체) / `txtPA_SPMON` / `ctxtSO_VKBUR-LOW`(센터, 필수) / `ctxtSO_KUNNR-LOW,HIGH` / `ctxtSO_WEEK-*`. 060 은 `txtSO_SPMON-LOW`. 058/060 은 **ZCRM_CICSVC_NEW** 를 읽는다(052/171/177 은 ZCRM_CICSVC).
- **정산 시점 규칙(2026-09-21 업무 확인)**: 월 마감 이후에는 그 기간월로 추가 접수를 받지 않는다(SAP 유입 상담건은 결제 완료분이고, 추가 접수는 새 상담번호로 들어와 다음 달 정산). 본사확정·전표는 마감 후 수행. → v0.01 의 FAIL P-01(전표 완료 후 추가 접수분 미정산)은 **철회**, E-22 PASS. 남는 유의점: 프로그램에 기간월 마감 전 확정을 막는 장치가 없어 월 중 조기 확정 시 화면 집계(ZCRM_RECPT 실시간)와 ZCRM_RECPTF(확정 스냅샷)가 어긋난다.
- E2E 데이터: 상담번호 2026091993001~005 / 접수 2026091900001~004 / FI 전표 3100069937(100047 구독)·3100069938(100032 일반). 결과서 `results\SAP_KR_단위테스트케이스및결과서_SR03000035400_전체프로세스통합_v0.02_20260921.xlsx` (24건 전건 PASS, 증적 22장, 시트에 '프로세스흐름' 포함; 개정이력 0.01~0.02). 생성 스크립트 `mk177e2e.ps1`, `make_fill_data_e2e.py`, `add_process_sheets.ps1`.

## 21. 미결 항목 체크리스트

- 업무·회계 결정이 필요해 테스트로 판정할 수 없는 항목은 `OPEN_ITEMS.md` 에 모은다. SR03000035400 작업을 다시 할 때는 이 파일을 먼저 읽고 `미결` 행을 사용자에게 다시 확인한다.
- 현재 미결: O-2 `ZCRM_KOSTL_T` 의 빈 `Z01·0520` 행 / O-3 기간월 마감 전 본사확정 통제 / O-4 Z11 설정 정리(test_ 문구, 예산센터, 1380). O-1 은 2026-09-28 확정(구독만 240074).

## 22. 2026-09-21~22 추가 개선 확인 결과

- **소비자가(ZPA2)**: ZCRMR052T01 141~142 상수, ZCRMR052F01 654~659 `구독 AND SUBSWRTMONTH > 0 → ZPA2` 분기. 30031356A 기준 PR00 39,500 / ZPA2 71,000. 조건레코드는 A911(VKORG ASKR/VTWEG 10) ⨝ KONP, KONP-KBETR 는 KRW 이라 395.00 = 39,500원.
- **Z02 자동판정**: ZCRMR052F01 2288~2298, 구독 모드에서 회수가능제조일(PRDRETDATE) > 부품제조일(PRDDATE_PRT) 이면 NCF='N', NCF_RSN_V='Z02'. 회수가능제조일은 N년이후/이내에 따라 다르게 계산됨.
- **센터 0520 = 구독사업팀**(TVKBT). ZCRM171S 정산 화면 SO_VKBUR 로 조회. 정산 ALV 맨 앞에 PRICETYPET(구독여부) 추가.
- **용역매출원가 계정**: ZCRMR171 93~113 `GV_SUBS_MODE='X'` → 정산유형 Z11, 아니면 Z08(ZCRMR0171_1T01 179~181 GC_STRART_*). FM `ZCRMF03_AS_SETTLE_FIPOSTING`(LZCRMG03U03)은 ZCRM_FIPOST_T(키 STRART+BUKRS+BUPLA+SECCO → HKONT/NEWKO/전기키/세금/SGTXT)와 ZCRM_KOSTL_T(키 STRART+BUKRS+VKBUR → KOSTL/ZZKOSTL)를 읽기만 한다. Z11/1100 NEWKO = **4305100 용역매출원가**(계정과목표 1000). 전표취소 `ZCRMF03_AS_SETTLE_CANCEL` 은 I_BUKRS·I_BELNR 만 받아 정산유형과 무관.
- 계정명은 SKAT 을 SE16N 조건 `G/L 계정` / `내역=*용역*` 으로 찾는다. 회사코드 레벨 개설은 SKB1(`G/L 계정`, `회사 코드`).
- 결과서 `results\SAP_KR_단위테스트케이스및결과서_SR03000035400_추가개선_회귀_v0.01_20260922.xlsx` (19건 PASS 18 / FAIL 1 = D-05 성능), 열람용 사본 `results\열람용\`. 최종 전표 3100069943(구독 4305100)·3100069942(일반 5328210).
- ZCRM060S/A 결과화면은 2026-09-22 에 HardCopy 가 -2147418113 으로 계속 실패(S20_NOMAX=1 로도). 값 확인은 DUMP 로 한다.

## 23. 2026-09-28 요구 정정 반영 확인 (결과서 추가개선_회귀 v0.02)

- **개선의뢰서 정정**: 구독 FI 계정은 용역매출원가(4305100)가 **아니라** 기존 A/S부품비 5328210. ZCRM171S 는 0520 전용이 아니라 모든 센터 조회. §22 의 4305100·0520 서술은 폐기.
- 구현은 **설정**: ZCRMR171 의 Z11(구독)/Z08(일반) 분기는 그대로, `ZCRM_FIPOST_T` Z11 NEWKO → 5328210, `ZCRM_KOSTL_T` Z11 × 0520·CT01~CT24 → KOSTL·ZZKOSTL 240074. 일반 Z08 은 대리점 센터(예 CT03 → 260011, 예산 260002).
- ZCRMR052: 선택화면 라디오 숨김 로직 주석(F01 1852~1867) → 052S 에도 센터조치/서비스점조치 둘 다. 부적합 자동판정 Z02 → **Z16 구독 기간 만료**(F01 2324~2334, 구독 모드만). 비구독은 회수가능제조일 > 부품제조일이어도 부적합 없음.
- 구독 문구 년수: FORM set_subs_pricetypet(F01 2739~2778), 구독기간 = WRTDATE − ISTDATE 개월 → 12의 배수면 `n년`, 아니면 `n.n년`(예 15개월 1.3년, 3개월 0.3년), 0 이하 `N년`. **ZCRMR171(TEXT-U03/U04)·ZCRMR177_1 은 아직 'N년' 고정.**
- 052S 조회 ALV 에서 SUBS_TYPE(구독) 컬럼이 빠지고 서비스 처리결과(PRICETYPET)가 첫 열.
- 데이터: 상담 2026092890001~003(대리점 100050, 센터 CT03, 생성자 ZTEST0928), 접수 2026092800001(71,000). 전표 3100069944(100050 구독)·3100069945(100034 구독 재전기), 3100069943 → 역분개 9930001228. 스크립트 `mk0928data.ps1`, `make_fill_data_0928.py`, `add_testdata_sheet0928.ps1`. S19 허용 접두어에 202609289 추가.
- 결과서 `results\SAP_KR_단위테스트케이스및결과서_SR03000035400_추가개선_회귀_v0.02_20260928.xlsx` (23건 PASS 22 / FAIL 1 = D-05 성능), 열람용 사본 `results\열람용\…_v0.02_열람용.xlsx`.
- 요청자 안내서 2026-09-28 갱신: artifact GdMM6P6bzpj6ei6kedK1ZT Version 3, Word `results\SR03000035400_구독정산_개선기능안내_20260928.docx` (scratchpad build_guide_docx_v3.js / guide_body_v3.html). 09-22 판 docx 는 구버전. 이 PC 에 LibreOffice 가 없어져 docx 확인은 Word COM ExportAsFixedFormat(17) 으로 한다.
- **v0.03 (2026-09-28 밤)**: ZCRMR171 구독 문구 수정 확인 - 조회(상세)는 실제 년수(FORM ZF_SET_SUBS_TEXT 에 ISTDATE·WRTDATE 인자, 1F01 1682~1734), 정산 줄은 `TEXT-U05` '구독무상' 고정(2F01 2380~2386). 텍스트 기호 U03·U04 삭제됨. 2F01 2440~2445 에 날짜 '00000000' PERFORM 이 남아 있으나 CLEAR 직전이라 무영향(정리 권장).
  새 대리점 **100068 경동성남동부대리점**(KNVV CT13 정상점, LFA1 123068) 상담 2026092891001~005(구독기간 12개월 단위: 36/24/60/12개월), 접수 2026092800002~005, 전표 3100069946(구독 5328210/240074)·3100069947(일반 5328210/260037). 결과서 `…추가개선_회귀_v0.03_20260928.xlsx`(28건 PASS 27 / FAIL 1 D-05), 스크립트 `mk0928b.ps1`, `make_fill_data_0928b.py`, `add_testdata_sheet0928b.ps1`.
  대리점 고르기: KNVV `고객=100030..100120`(범위 `..`, 조건은 인자 하나씩) + KNA1 NAME2 에 `(종료점)` 이 붙은 곳은 AUFSD 가 공백이어도 제외. S18 은 이관(STMS_IMPORT) 중인 세션을 만나면 ABORT 하니 다시 돌린다.
- **ZCRM_FIPOST_V 의 '정산유형명'** 은 텍스트 테이블 `ZCRM_STRART_T`(SPRAS·STRART·STRART_T, SM30 유지보수 있음, 패키지 Z33CRM) 값이라 뷰에서는 못 고친다. SE16N 조건 이름을 모를 때는 `S25_se16n_seldesc.vbs <테이블>` 로 조건 설명 목록을 뽑는다(DD27S=`뷰이름`, TVDIR=`뷰/테이블`).
- **ZCRM177A/AS 문서처리 실측 (2026-09-29, 사용자가 화면에서 체크박스+버튼 실행)**: 상담 2026092990001(구독)/002(일반), 센터 CT08 → 무상오더 4293369/4293370, 납품 85857225/226, 자재문서(953) 4921917627/628, FI 5300219255/256. 둘 다 5328130 아프터서비스비-부품A/S비(물류), 코스트센터 260014(CT08), 예산센터 공백 → 구독·일반 차이 없음(OPEN_ITEMS O-6). 저장위치 = 상담건 CENTER_CD 라 그 센터에 재고가 없으면 오더만 생기고 납품이 막힌다(2026091990008 / CT04 / 오더 4293368 잔존). 출고 FI 는 BKPF `참조 키`=자재문서번호+연도 로 찾는다. 테스트 데이터 스크립트 `mk0929_177.ps1`, S19 허용 접두어 202609299 추가.
- **통합 진행 확인 (2026-09-29 밤, 대리점 100033 코지하우징 / CT01 강북센터 / 상담 2026092990011 구독3년이후·012 구독2년이내·013 일반, 스크립트 `mk0929_e2e.ps1 -Part svc|new|req|m|sub|cnfrm`)**: 같은 상담건을 ZCRM_CICSVC·ZCRM_CICSVC_NEW·ZCRM_REQ·ZCRM_CICCOST_M/SUB/ADD 에 넣으면 177→052→171→058→060 이 모두 진행됨.
  177 오더 4293412~414 / 납품 85857279~281 / 출고 FI 5300219338~340(260006). 052 접수 2026092900001~003(71,000/39,500/39,500). 171 FI 3100069949(구독 240074/240074)·3100069950(일반 260006/260002). 058 FI 3100069951(구독)·3100069952(일반) 둘 다 5328170 / 260006 / 260002. 060S 2건 10,000, 060A 1건 5,000.
  공용 FM 은 **ZCRMF03_GET_SUBS_TEXT**(기능그룹 ZCRMG03, LZCRMG03U22) - 052F01 2749 / 0171_1F01 1697 / 177_1F01 3175 에서 호출. FM 없이 177 만 먼저 들어왔을 때 CALL_FUNCTION_NOT_FOUND 덤프 발생(이관 순서 주의).
  **ZCRM058 처리 순서**: 058S(구독) = CIC 대리점확정일 `ZCRM_CHRG_CNFRM`(CSLTNO·SPMON·CFDATE1) 필요 → btn[48] 구독 무상 기술료/출장비 확정 → btn[46] 전표(예 → 달력 오늘). 058A(일반) = btn[48] 무상지원금확정 → btn[2] 대리점 확정(예 → 달력) → btn[46] 전표(예 → 달력). 058A 에서 대리점 확정을 먼저 누르면 E "확정 되지 않았거나…". S19 허용 테이블에 ZCRM_CHRG_CNFRM 추가.
  SAP 연결 확인이 수 분 이상 멈추면 SAP GUI 스크립팅 접근 확인 팝업을 사용자가 눌러야 한다(대신 누르지 않는다).
- **결과서 v0.05 (2026-09-30)**: `…추가개선_회귀_v0.05_20260930.xlsx` = v0.04 시트 그대로 + 새 시트 **'통합테스트케이스'**(프로그램별테스트케이스 시트를 복사한 같은 양식, I-01~I-12 전건 PASS) + **'통합테스트증적(캡처)'**(30장). 생성 순서 `make_fill_data_0930.py` → `fill_template_subs.ps1` → `add_testdata_sheet0929.ps1` → **`add_integ_sheets.ps1 -Path -Json`**(payload 의 `integ` 키). 열람용 사본은 `build_plain_xlsx.py` 가 `integ.evidence` 도 그림으로 넣도록 수정됨.
- **상담건 센터 vs 대리점 소속 센터 (2026-09-30, KDQ)**: 2026.07.01~07 대리점 조치 8,687건 중 99.5% 동일, 서로 다른 값 0건, 상담건 센터 공백 42건(통신·도어락 계열 6곳, 기사명도 공백이라 058 대상 아님). **ZCRM058/060 은 `knvv~vkbur IN so_vkbur`(ZCRMR060_1F01 GET_BASE 75~104, INNER JOIN KNVV ASKR)로 대리점 소속 센터 기준**, ZCRM171 은 `ZCRM_CICSVC~VKBUR`(상담건 센터) 기준. 테스트 상담 2026093090001(VKBUR 공백)·002(CT13), 대리점 100068 → 058A CT13 2건 10,000 / 060A 2건 → 공백 센터도 집계됨 확인(미확정 상태로 남김). 스크립트 `mk0930_vkbur.ps1`, `S25_se16n_seldesc.vbs`.

## 24. 2026-09-30 변경분 통합 테스트 (결과서 추가개선_회귀 v0.06, 2026-10-01)

- **설정**: ZCRM_FIPOST_T·ZCRM_STRART_T 에 구독 전용 **S01~S10**(Z01~Z10 의 [구독] 복사본) 추가, Z11 삭제. ZCRM_KOSTL_T 는 **S01·S09 만**(코스트센터 = 센터별, 예산센터 = 240074). SE16N 조건 이름 `정산유형`.
- **로직**: ZCRMR171 구독 → S08(T01 181~182), 요청일자 기본값 이번 달(ZF_INITIALIZATION). ZCRMR058_1 구독 → S01(T01 123~124). ZCRMR177_1 ZF_CHK_DUP_CSLTNO(혼재 판정 성능 개선) + 구독 오더 ZCRM_ORDER_AUTO 프로그램명 ZCRM177AS.
- **결함**: D-06 171S 구독 전표 S046(S08 코스트센터 행 없음) / D-07 177A·177AS 구독·일반 분리 사라짐(GR_SUBS 미사용) / D-08 ZCRMR058_1F01 436·558·653·827·876 의 DELETE/UPDATE 가 SUBS_TYPE 조건 없음 → 058A 전표·확정이 구독 행에도 기록 / E-01 058S 구독 전표 예산 부족(240074 × 5328150).
- 테스트 상담 2026093090011~020(대리점 100184 CT18·100250 CT11·100215 CT08·100085 CT19), 접수 2026100100001~010(적합판정일 10.01 → 171 기간월 2026.10), 전표 3100069953~956. 사용자용 2026093090021~030(대리점 100234·100279·100256·100095·100005·100157·100200·100133) 미처리.
- 스크립트 `mk0930_e2e.ps1 -Set T|U -Part svc|new|req|m|sub|cnfrm`(**PowerShell 변수는 대소문자 구분 없음: 반복 변수 $c 가 공통값 $C 를 덮어써 템플릿 값으로 들어간 사고** → `$COMMON` 으로 변경), `S26_se16n_update.vbs`(테스트 상담번호 1행 필드 수정, 저장 modified=1 검증), `diff_scan.py`(S18 스캔 로그 비교), `make_fill_data_1001.py`, `add_testdata_sheet1001.ps1`, `add_integ_sheets.ps1 -Key integ2 -After <시트>`(시트 쌍 여러 개).
- 058/060 선택화면 센터는 범위(HIGH) 입력 칸이 없다 → 센터별로 조회. 058·171 전표생성 오류는 '정보: 에러 데이타가 있습니다' 팝업 뒤 ALV Message(ZMSG) 컬럼에서 읽는다(POPKEY:0 후 DUMP).
- **QA 예산 편성 (E-01, 2026-10-01)**: ZCO012(Z33COR012) 예산변경요청 → ZCO013(Z33COR013) 반영. ZCO013 은 상태 C(생성)·T(GW전송) 문서를 보여 주고 `Z33COF_BUDGET_REQUEST_APPLY` 를 `i_tapp='X'` 로 호출(그룹웨어 승인 불필요) → BAPI_COSTACTPLN_POSTPRIMCOST 로 버전 **B1** 에 증액, 상태 B. 승인자 사번은 SU3 `%LOGON_PERNO%`.
  ZCO012 화면: `ctxtP_BUKRS` `txtP_GJAHR` `radP_CK01`(예산센터) `cmbP_BGDTY`(Key A=추경, S20 `COMBO:`) → F8 → `txtGV_LTEXT`(적요), 그리드 `cntlCON1` 컬럼 KOSTL/KSTAR/PERIO/WTGBTR, Enter 하면 기초·전용·추경·소계·집행·가능액 표시. 통제기간 **3개월(분기)**. **저장(btn[11])은 확인창 없이 바로 저장**되고 GW 인터페이스(Z33COF_IF_BUDGET_SAP_TO_GW2)를 부른다.
  5328210 은 "예산 통제 계정이 아닙니다" → 171S 구독 전표는 예산 점검 대상 아님. 5328150·5328170 은 통제 대상. Z33COT007-WTGBTR 은 KRW 내부값(SE16N 10,000.00 = 1,000,000원).
  요청 문서: 40051(사용자, 240074×5328150 9·10월 각 1,000,000원) / 40052(Claude, 10·11·12월 각 1,000,000원, 적요 ZTEST_20261001…). 40051 은 사용자가 21:17 ZCO013 반영(상태 B, COSP KS10000000240074 B1 5328150 WKG009·010 각 1,000,000원). 40052 는 상태 C(ZCO013 승인은 권한 거부로 사용자 몫). ZCO012 화면 추경 칸은 월과 무관하게 1,000,000, 소계·가능액 0 으로 표시돼 신뢰 어려움 → 실제 반영은 COSP(오브젝트 번호 KS+KOKRS+KOSTL 10자리) 로 확인. ZCO013 결과화면 ALV wnd[0]/shellcont/shell, 버튼 tbar[1]/btn[17] 승인·btn[18] 취소.
- **E-01 해결 (2026-10-01 밤)**: 40051 반영만으로 058S 구독 전표 생성 성공(사용자 수행) - 100184 S 3100069959(CFDATE1 공백 유지 = D-08 수정 반영), 100184 일반 재전기 3100069958. 40052 는 상태 C 그대로(불필요). 100085 S 행 CFDATE1 2026.10.01 은 D-08 오염 잔존.
- **결함 재테스트 (2026-10-01 밤, 결과서 v0.07)**: D-06 해결(KOSTL_T S08 × 1100 25행, 171S 전표 3100069960~963 = 5328210/센터 코스트센터/240074) · D-08 해결(058_1F01 235·449·578·681·862·919 subs_type, 100184 로 058A 전표삭제·확정취소·무상지원금취소·재확정·대리점확정·전표 3100069964, 058S 전표삭제·전표 3100069965 - 반대쪽 행 불변) · E-01 해결 · D-05 해결(177A 21,574건 65초, 이전 95초).
  **D-07 미해결**: ZCRMR177_1F01 152~156 DELETE 는 들어왔지만 GV_SUBS_MODE 를 설정하는 곳이 없음(ZF_SET_SUBS_MODE 는 9/30 에 이미 없어짐) → 177AS·177A 둘 다 일반 건만. 조치안 = 152행 `IF sy-tcode = gc_tcode_subs.` 또는 START-OF-SELECTION 에서 gv_subs_mode 설정.
  058S 구독 취소(btn[47])는 CIC 확정일(ZCRM_CHRG_CNFRM)이 있으면 E "대리점 확정이 된 데이터" 로 막힌다(설계). 058S 버튼: 45 전표삭제 / 46 전표생성 / 47 구독 취소 / 48 구독 확정 / 40 지급대상데이터 / 43 상세 / 44 회계전표출력.
  결과서 `…추가개선_회귀_v0.07_20261001.xlsx` = v0.06 + 시트 '결함재테스트(1001)' K-01~16(PASS 14/FAIL 2) + '결함재테스트증적(1001)' 20장, FAIL내역 제목에 [해결]/[미해결]. 생성 `make_fill_data_1001b.py` → fill_template_subs → add_testdata_sheet1001 → add_integ_sheets integ / integ2 -After 통합테스트케이스 / integ3 -After 통합테스트케이스(0930). 열람용 v0.07.
  밤 22시 이후 HardCopy 가 새 세션에서도 -2147418113(restore 후 재시도도 실패, 화면 잠금 추정) → KOSTL_T S08·COSP·Z33COT006 캡처는 빠짐. 100085 S 행 CFDATE1 2026.10.01 오염값 잔존.
- **D-07 해결 (2026-10-02, 결과서 v0.08)**: ZCRMR177_1 70행 `gv_subs_mode = COND #( WHEN sy-tcode = gc_tcode_subs THEN 'X' ).` (KDD → KDQ 이관). 2026.09.30 완료건 포함 177AS 구독 16건 / 177A 일반 4건, 2026.09.19 177AS 7건(혼재 2026091990007 메시지) / 177A 3건. 결과서 `…추가개선_회귀_v0.08_20261002.xlsx` = 결함재테스트(1001) K-01~17 전건 PASS(K-17 혼재 추가), FAIL내역 전부 [해결], 생성 `make_fill_data_1002.py`. 재부팅 후 HardCopy 정상 복구 → 10-01 누락 캡처 3장 추가.
  KDD ABAP 편집기에서 SAPLOLEA MESSAGE_TYPE_X(BC-FES-CTL) 덤프는 같은 시간대 HardCopy 실패와 함께 생긴 SAP GUI 프론트엔드 상태 문제였고 PC 재부팅으로 해결됨.

## 25. 결과서 v1.00 - 현재 설정 기준 전체 프로세스 재테스트 (2026-10-02)

- 새 파일 `results\SAP_KR_단위테스트케이스및결과서_SR03000035400_구독대리점비용정산_v1.00_20261002.xlsx`(개정이력 1.00, 이전 이력은 추가개선_회귀 v0.08 참조) + 열람용. 시트: 프로그램별테스트케이스 E-01~E-26(프로세스 순), **프로세스흐름**(`add_process_sheet_v1.ps1`), 테스트데이터(`add_testdata_sheet_v1.ps1 -Doc177`), 테스트증적 48장. 페이로드 `make_fill_data_v1.py --doc177 results\_doc177_v1.json`.
- 데이터: 상담 2026100290001~010(대리점 100187 CT18·100216 CT08·100241 CT14·100282 CT23, 이전 미사용 정상점, LFA1 123xxx), 조치일자 2026.10.02, 생성자 ZTEST1002E, `mk1002_e2e.ps1 -Set V`. S19 허용 접두어 202610029 추가. 접수 2026100200001~010(71,000/39,500/Z16). 전표: 171S 3100069966~968(5328210/센터/240074) · 171 969~971(…/260002) · 058S 972~975(5328170/센터/240074) · 058A 976~978(…/260002) · 출고 FI 5300219397~406(5328130/센터/공백). 177 오더 4293442~451, 납품 85857304~313, 자재문서 4921917807~816(사용자 수행).
- **신규 결함 D-09**: ZCRM_ORDER_AUTO PROGRAM_NM 이 구독 건도 'ZCRM177A'(F01 960·1205·1339 `IF sy-tcode = gc_tcode_subs` 가 문서 생성 시점에 안 맞음, 09-29 오더도 동일). 조치안 = gv_subs_mode 로 판정. 정산·전표 영향 없음.
- **S19 템플릿 경로 사고**: 파이썬으로 ps1 을 만들 때 `results
aw` 의 `
` 이 CR 로 들어가 템플릿을 못 찾았고, S19 는 "template tsv missing" 을 로그만 남기고 지정 필드만 채운 **불완전한 행을 저장**한다(exit 는 0). 삽입 뒤 반드시 템플릿 컬럼(VKORG 등)이 채워졌는지 확인. 정리는 `S27_se16n_delete.vbs`(접두 202610029 + ERNAM 검증, 1~3행, &LOCAL&DELETE_ROW, 저장 팝업 deleted=n 확인). PowerShell 에서 삭제 op 토큰은 `'d'+'el'` 로 런타임 조립(하네스 차단 회피).
- 060 선택화면에는 radPA_NEW 가 없다(txtSO_SPMON-LOW·ctxtSO_VKBUR-LOW·ctxtSO_KUNNR-LOW 만). 171·058 전표는 10-02 에 모두 정상(예산 240074×5328150 10월 1,000,000 범위 내).

## 26. CTS 이관 로그 읽기 / D-09 진짜 원인 (2026-10-02 오후)

- **D-09 진짜 원인**: `ZCRM_ORDER_AUTO-PROGRAM_NM` 이 **C(8)** 이라 'ZCRM177AS'(9자)가 'ZCRM177A' 로 절단(사용자 디버깅으로 확인, F01 960 분기는 정상 진입). gv_subs_mode 수정은 KDQ 반영됐지만 길이 때문에 효과 없음. 조치안 A = 8자 값(예 'ZCRM177S') 상수로 저장 / B = 데이터엘리먼트 길이 확장(DDIC). 2026100290013(100085 CT19, 미처리)으로 재확인 예정. 11·12 는 수정 전 처리분.
- **이관 로그 읽기(스크립트)**: SE01 → `SELECT:tabsMAINTABSTRIP/tabpTSSN` → `SET:tabsMAINTABSTRIP/tabpTSSN/ssubCOMMONSUBSCREEN:RDDM0001:0210/ctxtTRDYSE01SN-TR_TRKORR=<요청>` → `PRESS:…/btn%#AUTOTEXT032`(Logs) → 개요 리스트(SAPMSSY0/120, lbl[19,y] 단계명·lbl[84,y] RC) → `FOCUS:lbl[19,y]` + `KEY:2` 로 상세 로그 → `BTN:16`(전체 펼치기 Shift+F4) → `LIST`+`KEY:82` 반복. S20 에 `FOCUS:` 스텝 추가. 로그 파일명 = 원본SID + 단계문자(E export/A DDIC활성화/H import/G 생성) + 요청뒷자리 + .대상SID. `\erpdev\sapmnt	rans\log` 는 PC 에서 직접 못 연다.
- **KDDK9A01ZP 노란불 원인**: KDD Export RC4 = 요청에 `LIMU FUNC Z33CRM`(패키지명을 FM 으로 잘못 등록) → "function module Z33CRM does not exist - ignored". KDQ DDIC 활성화 RC4 = 뷰 ZCRM_KOSTL_V "W- 키필드가 시작 뷰에 같이 없습니다"(텍스트 조인 ZCRM_STRART_T·TVKBT 의 SPRAS 키 미포함; 기본 테이블 키는 모두 포함) + 종속 ZCRM_KOSTL_T. 둘 다 기능 영향 없음. 안내 문서 `results\CTS_이관_노란불_원인확인_및_초록불_만들기_20261002.docx`.
- **S20 세션 선택 수정**: 생성 전 세션 Id 목록을 기억해 두고 새로 생긴 세션을 고른다(ElementAt(Count-1) 이 사용자 SE01 세션을 집어 ABORT 14 났던 문제). SE16N 라벨: E071 `요청/태스크`·`오브젝트이름`, DD27S `뷰이름`, REPOSRC `프로그램이름`(UDAT/UTIME = 이관 반영 시각).
- **PowerShell 1건 배열 사고**: `@( @('a','b') )` 와 `if` 식 대입은 1건 배열을 풀어 버린다 → `@( ,@(...) )` + 직접 대입. mk1002_e2e.ps1 에 행 길이 가드 추가. 잘못 들어간 12자리 상담번호 행 10개는 S27(12·13자리 허용)로 삭제.
- **D-09 해결 확인 (2026-10-02 23시, 결과서 v1.01)**: B안 반영 - ZCRM_ORDER_AUTO-PROGRAM_NM CHAR 8 → 10(DD03L LENG 10, 조건 라벨 `테이블이름`), F01 22:17 재이관. 재고 선점 오더 4293457·4293458(B안 이관 전 013 시도분, ORDER_AUTO 기록 없음)을 VA02 로 삭제(S20 `MENU:wnd[0]/mbar/menu[0]/menu[10]`=판매문서→삭제, 팝업 `usr/btnSPOP-OPTION1`). 013 재처리 → 오더 4293459(미결 잔존)·4293460 / 납품 85857316 / 자재문서 4921917822 / FI 5300219411(5328130/260037) 모두 PROGRAM_NM ZCRM177AS. 결과서 `…구독대리점비용정산_v1.01_20261002.xlsx`(E-05 PASS, FAIL 0, 개정이력 1.01) = `make_fill_data_v1.py --doc177 results\_doc177_v101.json`(fail_rows·revision·out 키 지원).
- S20 추가: `MENU:<id>`(메뉴 select), `FOCUS:<id>`; 새 세션이 S000(로그온 화면)이면 20초까지 대기. 디버거(Exclusive Mode)가 열려 있거나 접속이 로그오프되면 스크립트가 attach 단계에서 멈춘다 → 사용자가 디버거 종료/재로그온.
- ATP: 177 납품 실패 "재고 없음"은 MARD 재고가 있어도 VBBE(미결 요구량)에 선점 오더가 있으면 난다. 조건 라벨 VBBE `자재번호`·`플랜트`, VBAK `판매 문서`.
- **기술료·출장비(ZCRM_CICCOST_M-MAXTECHPRICE/TRIPPRICE) 실측 (2026-10-02 밤)**: CIC 적재 FM(LZCRMG02U94, `C.MAXTechPrice/100`, `C.TripPrice/100`)이 채우고, ZCRMR058_1 은 ZF_BUILD_COST_SUM/ZF_APPLY_COST(1542~1590)로 대리점별 합산해 **표시만** 한다(058S·060S 컬럼). 확정(367~394 `lt_assptfeeh1-assptfee = gt_data_total-assptfee`)과 전표(530~548 `lv_wrbtr = assptfee`)는 건수×단가(night+deepnight+weekday+holiday, 060F01 1330)만 쓴다. ZCRM_ASSPTFEEH1 에 기술료·출장비 컬럼 없음. 실측: 013 에 기술료 10,000/출장비 20,000(S26, 접두 202610029 허용 추가) → 058S 표시 10,000/20,000, ASSPTFEE 5,000, 확정·전표 3100069979 = 5,000. → **결함 후보 D-10**(요구사항 확인 필요). 실데이터 예: 출장비 20,000 고정, 기술료 5,000~26,000, CT99(나이스엔지니어링) 건.
- **D-10 B안 확정 (2026-10-03)**: 구독 지원금 = 기술료 + 출장비(건수 단가 제외). 수정안 `results\D-10_B안_수정안_기술료출장비_20261003.txt`(DDIC ASSPTFEEH1 +MAXTECHPRICE/TRIPPRICE, 060F01 1532~1553 금액 대체, 058F01 ZF_APPLY_COST CHECK red + 합산, lt_feeh 두 필드, SET_MCOMP 저장·0원 차단). 기존 청구 프로그램 ZCRM185(ZCRMR185 온수매트 수리 비용처리)는 SUB 92000000(기술료)/92000001(출장비) 행을 판매오더 품목(ZPR2)으로 유상 청구 - 058S 는 FI 지급이라 원천만 공유. 060 구조: line_idx 1 건수 행 / 2 기사별 금액 행 / 999 합계 행; ASSPTFEEITM 은 확정에서 미사용.
- **사용자 지시**: 이후 테스트 문서(결과서·설명서·FAIL내역)에는 **기존 로직 vs 변경 로직 비교표**를 반드시 넣는다(메모리 test-doc-logic-diff).

## 27. D-10 B안 / D-11 재테스트 (2026-10-03 새벽, 결과서 v1.02)

- **D-10 B안 반영 확인(KDD → KDQ 02:09)**: ZCRM_ASSPTFEEH1 +MAXTECHPRICE·TRIPPRICE(CURR 13,2, 참조 ZCRM_CICCOST_M-WAERS) / ZCRMR060_1F01 1533~1567(금액 행 line_idx 2·합계 행 999 = 기술료+출장비, 건수 단가 CLEAR) / ZCRMR058_1F01 ZF_APPLY_COST 1606~1625(빨강만 재계산), ELSE 113~116(저장값 표시), SET_MCOMP 356~363(0원 제외)·392~397(저장).
- **D-11**: 1차 수정은 CURR 내부값을 `TYPE i` 에 그대로 더해 300 표시. **CURR(참조 WAERS) 필드의 KRW 내부값은 1/100**(10,000원 = 100.00). ALV 는 set_currency_column/CFIELDNAME 으로, SE16N 은 참조 WAERS 가 같은 테이블일 때만 환산해 보여 준다. 정수 원 단위 필드에 넣을 때는 `BAPI_CURRENCY_CONV_TO_EXTERNAL`(FORM ZF_SUBS_FEE_TO_INT, 060F01 2166~2197) 로 환산. 058 전표 WRITE(554)는 정수 원 단위 전제.
- 재테스트 데이터: 013(100085/CT19, 기술료 10,000/출장비 20,000) 전표 3100069979 → 역분개 9930001233 → 신규 **3100069980(30,000)**. 058S 구독 취소(btn[47])는 ZCRM_CHRG_CNFRM 행이 있으면 막히므로 S27 로 지웠다가 `mk1002_e2e -Set W -Part cnfrm` 으로 복구. 0원 케이스 014(100005/CT02, `-Set W2` 6개 Part) 확정 → "0 건 저장", 헤더 행 없음(제외 메시지는 덮임 = D-12 권고). D-13 권고: 새 컬럼 참조통화가 ZCRM_CICCOST_M-WAERS 라 SE16N 에 100.00 으로 보임.
- 결과서 `…구독대리점비용정산_v1.02_20261003.xlsx` = 사용자 수정 v1.01 복사본에 시트 추가: '변경내역(기존vs변경)'(기존/변경 로직 비교표 8항목, `add_logic_diff_sheet.ps1`), '결함재테스트(1003)' R-01~R-18(PASS 17 / FAIL 1 = D-11 해결) + 증적(`add_integ_sheets -Key integ4 -After 테스트증적(캡처)`), FAIL내역 D-10·D-11·D-12·D-13 추가(`add_fail_rows.ps1` → `add_fail_pictures.ps1`), 테스트데이터 5절, 개정이력 1.02·표지(`finish_v102.ps1`). 페이로드 `make_fill_data_v102.py`(v101 페이로드 기반). 열람용은 `dump_xlsx.ps1` → `build_plain_xlsx.py`(integ4 추가).
- 쉬운설명 문서 10-03 판: scratchpad easy\easy_v102.html → Word COM 으로 docx. **10-02 판 docx 는 그림이 링크로만 들어가 있었음** → 10-03 판부터 InlineShapes 를 문서에 포함(embed)시켜 저장.
- S18(KDQ) 인자 순서는 `<tag> SCAN <문자열> <SET…>` (S18k 와 동일). mode 를 빼면 문자열이 mode 로 들어가 0건.
  (줄 번호는 D-11 반영 뒤 02:09 활성본 기준. D-13 의 '참조통화' 설명은 10-07 에 정정 - 아래 참조.)

- **10-06 보완**: `mk1002_e2e.ps1 -Set W3` = 상담 2026100290015(100005/CT02, 구독) 추가 후 S26 으로 기술료 7,000/출장비 20,000 → 058S 100005 가 014(0원)+015 합산 27,000(R-11), 060S 건수 2·27,000. 060A 100187 5,000 유지(R-12). ZCRM_CICSVC(_NEW) 의 VKORG 는 V 세트(001~)도 공백이라 정상(VKORG 는 CICCOST_M 에만). COSP 240074×5328170 Z1 두 행은 BEKNZ D(차변 909.07)/C(대변 181.81-) 행이지 특정 전표 귀속이 아니다. 10-02 판 쉬운설명 docx 도 그림 포함으로 재저장함.
- **10-07 전반 재검토(결과서 v1.02 최종, 파일명 _20261007)**: 결함재테스트(1003) R-01~R-18. 추가 확인 = 016(100005, 3,000/5,000 → 합산 35,000 증명) · R-13 확정 후 CIC 금액 변경(058S 저장값 유지, **060S 는 CIC 현재값 31,000 = D-15**) · R-14 radPA_TAB(확정 데이터) 모드 · R-15 btn[40] 지급대상데이터조회(화면 101)·btn[43] 상세리스트(SUBMIT ZCRMR060_1) · R-16 058A 100187 쓰기 회귀(3100069976 역분개 → 3100069982, 버튼 2 대리점확정/5 대리점확정취소/45/46/47 무상지원금취소/48) · R-17 100085 전표 재생성(3100069980 → 9930001235 → 3100069983) · R-18 혼합 선택 = 2026.09 CT14 에서 100234(set U, 0원) + 100241(새 상담 2026100290018, `-Set W6`, S26 로 INDATE·**CICSVC_NEW RSLTDATE** 2026.09.30 / 4,000·6,000, cnfrm 없음) → '1 건 저장' → 100241 만 헤더 → btn[47] 취소 성공. **058 의 기간월은 ZCRM_CICSVC_NEW-RSLTDATE 기준**(CICCOST_M INDATE 만 바꾸면 안 잡힘). CT02 에는 100005 외 정상 대리점이 없음(100590 은 동사무소) → `-Set W5` 는 미사용. 015·016 은 잠시 0 으로 바꿨다가 7,000/20,000·3,000/5,000 으로 원복.
  **D-14(기존 코드)**: SET_FIDEL(058F01 632~722)이 e_result 를 lv_check_error 로 받고 lv_e_result 를 검사 → 취소 실패해도 헤더 BELNR/BUDAT 삭제. **D-13 정정**: ASSPTFEE 는 참조 ZCRM_ASSPTFEEH-WAERS(다른 테이블)이고 30,000.00 은 TYPE i 를 그대로 저장한 원 단위 값, 새 컬럼은 내부값 → 한 행에 단위 혼재. ZCRM059·059A·059AS·059S 가 TSTC 에 있고 ZCRMR058_1 1498 이 ZCRM059S 를 구독 모드로 취급(미테스트). INSERT_DB_TABLE(060F01 1907~1942, ZCRM_ASSPTFEES1 월 DELETE)은 호출처 없음. CT02 정상 대리점(KNVV AUFSD 공백)은 100590(동사무소)·100879 뿐(100030~100400 은 전부 Z8).
  스크립트: add_integ_sheets.ps1 이 행별 `date` 우선, dump_xlsx.ps1 이 날짜 셀을 yyyy-MM-dd 로, add_logic_diff_sheet.ps1 테스트데이터 구간 재실행 안전, 한글 리터럴 ps1 은 UTF-8 BOM 으로 저장. **Bash 도구의 cwd 는 호출 사이에 바뀔 수 있어 상대경로 append 가 엉뚱한 폴더(results\CLAUDE.md, scripts\CLAUDE.md)에 파일을 만들었음 → 항상 절대경로**(메모리 bash-cwd-absolute-paths).
- **쉬운설명 docx 그림 (2026-10-07 사용자 지적)**: 10-02 판부터 scratchpad\easy 의 그림이 900x456 으로 잘린 사본이었고, Word 가 HTML 을 웹 레이아웃으로 열어 넓은 그림이 창 폭에서 더 잘려 보였음. 고친 방법 = scratchpad\easy_full 에 results\shots 원본(2560x1368)을 폭 1400 으로만 줄여(자르지 않음) 복사 → Word COM 으로 열어 View.Type = 3(인쇄 레이아웃), A4·여백 1.5cm, 모든 InlineShape 를 LockAspectRatio + Width = 사용 가능 폭(510pt) 으로 맞춘 뒤 docx 저장(3.4MB, 25쪽). 앞으로 docx 에 넣는 캡처는 **항상 전체 화면 원본**을 쓴다.
- **캡처 안에서 열이 잘릴 때 (2026-10-07)**: SE16N 결과처럼 ALV 가 창보다 넓으면 원본 캡처 자체에서 오른쪽 열이 잘린다. S12 는 S12_SHOT=1 일 때 요청 컬럼의 첫 번째로 FirstVisibleColumn 을 옮기는데 repaint 전에 HardCopy 를 찍어 효과가 없던 것을 WScript.Sleep 1000 추가로 고침 → 보고 싶은 열을 컬럼 목록 맨 앞에 두면 그 열부터 보이고 키 열(고정 열)은 그대로 남는다(ZCRM_RECPT 는 CSLTNO,...,NETWR 로 정가 열까지 캡처). BSEG 는 SE16N 대신 **FB03 데이터 입력 뷰**(S20 TC:FB03 → 	xtRF05L-BELNR·ctxtRF05L-BUKRS·	xtRF05L-GJAHR → KEY:0, 그리드 cntlCTRL_CONTAINERBSEG/shellcont/shell)가 계정·금액을 한 화면에 보여 줘 증적에 낫다(코스트센터 열은 화면 밖).

## 28. 10-07 업무 결정 반영 확인 (결과서 v1.03)

- **결정**: ① 기술료·출장비는 구독 N년이후(SUBSWRTMONTH > 0) 건만 → **D-16** ABAP(ZCRMR058_1F01 1577 / ZCRMR060_1F01 236 `DELETE lt_base WHERE subswrtmonth <= 0.`, KDQ 22:02). ② 구독 정산 전표 코스트센터 = 240074 → **D-17** 설정(ZCRM_KOSTL_T S01 50행·S08 25행 KOSTL 240074, ZZKOSTL 240074). ③ 177AS 출고 원가(5328130)는 범위 제외 → **D-18** 검토만(results\D-18_…txt: 오더 VBAK-KOSTL 공백·TVAUK ASKR 0건·OKB9 1100 기본 270006·ZGGB/ZX Exit 없음 → 출고 BAdI 추정, 안 A 주문사유+OVF3 / B 파생로직 분기 / C VBAK-KOSTL).
- 확인(수정확인(1007) R-19~R-24 전건 PASS): 019(100005, N년이내 SUBSWRTMONTH 0, 9,000/9,000, `-Set W7`) 추가해도 058S 35,000 유지 / 018 SUBSWRTMONTH 0 → 0원(원복 12) / 171S 100187 3100069966 역분개 → **3100069984**(5328210·240074/240074) / 058S 100085 3100069983 → **3100069985**(5328170·240074/240074) / 일반 171 3100069969 → 3100069986(260014/260002), 058A 3100069982 → 3100069987(260014/260002).
- 171S 결과 ALV 에는 CT18 의 다른 구독 대리점(100184)도 같이 나오므로 ROW0 가 아니라 **대리점 행의 BELNR** 을 읽어야 한다(ROW0 를 읽어 옛 전표 3100069961 을 검사했던 실수).
- 결과서 `…_v1.03_20261007.xlsx` = v1.02 사본 + '수정확인(1007)'·'수정확인증적(1007)'(`make_fill_data_v103.py` integ5, `add_integ_sheets -Key integ5 -After 결함재테스트증적(1003)`), FAIL내역 D-16·D-17 [해결]·D-18 [보류], 변경내역 9·10행, 테스트데이터 6절, 개정이력 1.03(`finish_v103.ps1 -FromRow 7 -ToRow 8`). 쉬운설명 docx 에 ⑤-2 절 추가.
- 미결로 남는 것: D-18(출고 원가), D-12~D-15 권고 반영, ZCRM059 계열 T-code, 운영 240074×5328150 예산, O-3·O-5.
- **FAIL내역 캡처 '셀에 배치' (2026-10-08 사용자 요청)**: `add_fail_pictures_incell.ps1` - 각 결함 행의 E열부터 `Range.InsertPictureInCell`(Excel 365 그림 셀 값)로 넣고 여러 장이면 F·G열. **Excel 이 Visible=True 이고 해당 시트가 활성·셀이 Select 된 상태에서만 실제로 들어간다**(숨김/비활성이면 오류 없이 빈 셀로 저장됨). 그림 셀은 COM 으로 Value2·Text 가 비어 읽히므로 확인은 PrintWindow 창 캡처로 한다(`scratchpad\fail_final2.png` 방식, Excel 창을 SetForegroundWindow 후 user32 PrintWindow). ScreenUpdating=False 면 HRESULT 오류. 열람용(openpyxl)은 같은 셀 위치에 떠 있는 그림으로 대신 넣는다(build_plain_xlsx.py).
- **증적 캡처 배치 규칙 (2026-10-08 사용자 지시)**: 결과서에 별도 '증적' 시트를 만들지 않는다. 캡처는 각 테스트케이스 시트의 해당 행 L열부터 '셀에 배치'(여러 장이면 M, N …), FAIL내역은 E·F·G열. 스크립트 `add_case_pictures_incell.ps1`·`add_fail_pictures_incell.ps1`(Excel Visible 필수). v1.03 에서 증적 시트 3개 삭제, 범례 문구를 'L열부터' 로 수정. 열람용은 build_plain_xlsx.py 가 같은 셀에 떠 있는 그림으로.
