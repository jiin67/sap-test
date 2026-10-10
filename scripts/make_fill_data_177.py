# -*- coding: utf-8 -*-
# make_fill_data_177.py : JSON payload for the unit-test result workbook of the 2026-09-19 scenario
#   "ZCRMR177_1 일시불구독 개선 - 일반/구독 조회 분리 (ZCRM177A / ZCRM177AS)". Consumed by fill_template_subs.ps1.
import io, json, os
from PIL import Image as PILImage

OUT = r"C:\sap-test\results\_fill_data_177.json"
SHOTS = r"C:\sap-test\results\shots"
THUMB = r"C:\sap-test\results\shots\_tpl_177"
os.makedirs(THUMB, exist_ok=True)
DATE = "2026-09-19"
missing = []


def thumb(fn, maxw=760):
    src = os.path.join(SHOTS, fn)
    if not os.path.exists(src):
        missing.append(fn)
        return None
    dst = os.path.join(THUMB, fn)
    im = PILImage.open(src)
    im.thumbnail((maxw, maxw))
    im.save(dst, optimize=True)
    return dst


P, F, N = "PASS", "FAIL", ""
BLK = "BLOCKED"
NAP = "해당없음"

# (id, 그룹, 케이스, 예상결과, 심각도, 결과, 테스트데이터/실제결과)
CASES = [
    ("P-01", "사전준비", "테스트 데이터 생성",
     "일반 유상/일반 무상/구독 N년이후/0경계/보증일공란/다품목(혼재·동일)/기간경계 데이터를 만든다",
     "Major", P,
     "SE16N 으로 ZCRM_CICSVC 14행(상담번호 2026091990001~013, 생성자 ZTEST0919, CLMART='A'), ZCRM_CICCOST_M 12행(VKORG ASKR, 센터 CT04, 조치일자 2026.09.19), "
     "ZCRM_CICCOST_SUB 12행(COST=0) 삽입. 무상부품(SUB, COST=0) 행이 없으면 GET_DATA 가 상담번호를 지우므로(ZCRMR177_1F01 165행) 전 건에 필요. 상세는 '테스트데이터' 시트"),

    ("TC-01", "활성화", "프로그램 문법검사 및 활성화",
     "Field name is too long 오류 없음, 신규 3필드 미정의 오류 없음, 전체 객체 활성",
     "Critical", P,
     "RS_ABAP_SOURCE_SCAN 으로 활성 소스 전량 덤프: ZCRMR177_1 외 INCLUDE 10개, 8,923행을 결번 없이 읽음. 두 T-code 실행 시 구문오류·덤프 없음. "
     "※SE38 편집기에서의 문법검사 자체는 수행하지 않음(조회 전용 원칙)"),
    ("TC-02", "T-Code", "신규 T-Code ZCRM177AS 실행",
     "T-Code 오류 없이 ZCRMR177_1 선택화면·구독 화면 진입, 권한/화면번호 오류 없음",
     "Critical", P,
     "TSTC: ZCRM177AS → 프로그램 ZCRMR177_1 / 화면 1000 (기존 ZCRM177A 와 동일). 실행 시 선택화면 정상, 결과화면 제목 '[구독] AS 비용처리'"),
    ("TC-03", "T-Code", "기존 T-Code ZCRM177A 실행 회귀",
     "기존과 동일하게 실행되고 신규 T-Code 추가 영향 없음",
     "Critical", P,
     "ZCRM177A 선택화면·결과화면(제목 'AS 비용처리') 정상. 선택조건 구성(전기일자/회사코드/판매조직/센터/상담번호/기사명/조치일자/완료건 포함) 변화 없음"),
    ("TC-04", "조회분리", "ZCRM177A 는 일반 건만 조회",
     "SUBS_TYPE 공란 건만 조회되고 S 건은 제외",
     "Critical", P,
     "조치일자 2026.09.19 → 4건(2026091990001·002·007·013) 모두 비구독. 구독 6건(003·004·005·006·008·012) 제외됨. "
     "소스: GR_SUBS 에 E/EQ/'S' 를 넣어 c~subs_type IN gr_subs 로 제외(ZCRMR177_1F01 3113~3116)"),
    ("TC-05", "조회분리", "ZCRM177AS 는 구독 건만 조회",
     "SUBS_TYPE='S' 건만 조회되고 공란 건은 제외",
     "Critical", P,
     "같은 조건 → 7건(003·004·005·006·007·008·012) 모두 구독. 비구독 4건 제외. 소스: I/EQ/'S'(3104~3107)"),
    ("TC-06", "조회조건", "조회 시작일/종료일 경계",
     "시작일·종료일 당일 데이터 포함, 기간 밖 제외",
     "Major", P,
     "조치일자 2026.09.17~09.19 → 시작일 건 2026091990009(09.17) 포함, 기간 밖 2026091990010(09.16) 제외. "
     "※조회조건은 시나리오에 적힌 RSLTDATE 가 아니라 ZCRM_CICCOST_M-INDATE('조치일자', SO_INDAT) 이다"),

    ("TC-07", "연동", "SUBS_TYPE 연동",
     "S 는 S, 일반은 공란으로 수신되고 이전 행 값이 잔존하지 않음",
     "Critical", P,
     "SAP 적재값(ZCRM_CICSVC) ↔ 화면 표시가 11건 전부 일치, 행 간 값 잔존 없음. "
     "※ZCRMR177_1 에는 Native SQL(EXEC SQL)이 없고 ZCRM_CICCOST_M ⨝ ZCRM_CICSVC(CLMART='A') Open SQL 조인이다. CIC 원천 DB 와의 직접 대조는 이 프로그램 범위 밖"),
    ("TC-08", "연동", "WRTDATE 연동", "원천 날짜값이 잘림·자리이동 없이 적재",
     "Major", P,
     "적재 FM ZCRMF02_RFC_ZCRM_CICSVC_ALL(기능그룹 ZCRMG02)의 Native SQL 에 WRTDATE 가 정상 추가되어 있다. "
     "LZCRMG02U73 252행 'A.WRTDATE' ↔ 314행 ':GT_ZCRM_CICSVC_NEW-WRTDATE', LZCRMG02U27 281행 'A.WrtDate' ↔ 356행 ':GT_ZCRM_CICSVC-WRTDATE'. "
     "ZCRM_CICSVC-WRTDATE 는 DATS(8) 이고 SAP 적재값 저장·조회 일치(2026091990003=2027.09.19, 006=공란). "
     "※화면 프로그램 ZCRMR177_1 은 WRTDATE 를 읽지 않는데, 구독여부 판정이 SUBSWRTMONTH 만으로 이루어지므로 요구사항상 문제 없음(요구사항 D는 적재 테이블 필드 추가, F는 TEXT 표시). "
     "원천 CIC 값과의 실데이터 대조는 동기화 실행이 필요"),
    ("TC-09", "연동", "SUBSWRTMONTH 양수 연동", "부호·숫자 동일, 형변환 오류·덤프 없음",
     "Critical", P, "2026091990003·012 SUBSWRTMONTH=12 → 화면 '구독무상(N년이후)'. 006=3 도 동일. 덤프 없음"),
    ("TC-10", "연동", "SUBSWRTMONTH 음수 연동", "음수 부호와 절대값 유지",
     "Critical", N,
     "※미실행(BLOCKED) - QA 에서 음수 데이터를 만들 수 없음. ZCRM_CICSVC-SUBSWRTMONTH 는 INT2(부호 있음)라 DDIC 상 음수 저장은 가능하지만, "
     "SE16N 편집으로 '-6' 과 '6-' 모두 오류 없이 저장되지 않고(변경 완료 팝업 없음), SM30 유지보수 다이얼로그도 없다(E: 유지보수 다이얼로그를 완료하지 않았거나 생성하지 않았습니다). "
     "소스상 분기는 'IF pv_subswrtmonth > 0 … ELSE'(F01 3139~3143)이므로 음수는 0 과 같은 ELSE 경로이며, 그 경로는 TC-18 로 확인함"),
    ("TC-11", "연동", "SUBSWRTMONTH 0 연동", "0 유지, 초기값 잔존 문제 없음",
     "Major", P, "2026091990005 SUBSWRTMONTH=0 저장·조회 일치"),
    ("TC-12", "연동", "다건 APPEND / 행별 값 매핑",
     "행별 값이 정확히 매핑되고 첫 행 값이 후속 행에 복사되지 않음, CLEAR 로 행 간 값이 섞이지 않음",
     "Critical", P,
     "적재 FM: LZCRMG02F01 APPEND_CMS_ZCRM_CICSVC_NEW 에서 237행 'GT_ZCRM_CICSVC_NEW-SUBSWRTMONTH = GV_SUBSWRTMONTH.' → 240행 APPEND → "
     "241행 'CLEAR: GT_ZCRM_CICSVC_NEW, GV_SUBSWRTMONTH.' 로 보조 변수까지 초기화하므로 행 간 잔존 없음. "
     "화면: ZCRM177AS 7행이 각각의 SUBSWRTMONTH(12/0/0/3/12/12/12)에 맞는 TEXT 로 표시되고, 행마다 CLEAR pv_subs_text(F01 3133) 후 판정"),
    ("TC-13", "연동", "SELECT/INTO 순서",
     "SELECT 와 INTO 개수·순서 일치, 기존 필드와 신규 필드가 뒤바뀌지 않음",
     "Critical", P,
     "적재 FM Native SQL 의 SELECT 목록과 INTO 목록 순서 일치를 소스로 확인: "
     "LZCRMG02U73 251~253 (A.SUBS_TYPE / A.WRTDATE / A.SUBSWRTMONTH) ↔ 313~315 (:…-SUBS_TYPE / :…-WRTDATE / :GV_SUBSWRTMONTH), "
     "LZCRMG02U27 280~282 ↔ 355~357. 신규 3필드는 기존 필드(CSYMP_T3 등) 뒤에 추가되어 기존 매핑이 밀리지 않는다. "
     "화면 프로그램의 SELECT 는 INTO CORRESPONDING FIELDS OF TABLE 이라 이름 매핑(F01 68~71, 94) → 순서 밀림 구조적으로 불가. 기존 필드 값 정상"),

    ("TC-14", "TEXT", "일반 유상 TEXT", "구독여부 = 유상, 내부코드 S 미표시",
     "Critical", P, "2026091990001 (ISFREE 공란) → '유상'"),
    ("TC-15", "TEXT", "일반 무상 TEXT", "구독여부 = 무상, 구독무상으로 오표시되지 않음",
     "Critical", P, "2026091990002·007·013 (ISFREE='Y', SUBS_TYPE 공란) → '무상'"),
    ("TC-16", "TEXT", "구독무상 N년 이후 TEXT", "구독여부 = 구독무상(N년이후)",
     "Critical", P, "2026091990003 (S, 12) → '구독무상(N년이후)'. 'N년이상'·'S' 표시 없음"),
    ("TC-17", "TEXT", "구독무상 N년 이내 TEXT (음수)", "구독여부 = 구독무상(N년이내)",
     "Critical", N, "※미실행(BLOCKED) - TC-10 과 같은 사유(음수 데이터 생성 불가). 동일 ELSE 분기는 TC-18 로 확인"),
    ("TC-18", "TEXT", "구독무상 0 경계 TEXT", "구독여부 = 구독무상(N년이내)",
     "Major", P, "2026091990005 (S, 0) → '구독무상(N년이내)'. 소스 분기 기준 '> 0' 이 아니면 이내"),
    ("TC-19", "화면", "내부코드 비노출",
     "구독여부 컬럼에 업무 TEXT 표시, SUBS_TYPE 'S' 가 사용자 컬럼으로 직접 표시되지 않음, 제목은 '구독여부'",
     "Critical", P,
     "필드카탈로그에 추가된 것은 SUBS_TEXT 뿐이고(C01 979~987) SUBS_TYPE·SUBSWRTMONTH 는 표시열·숨긴열 어디에도 없다. 컬럼 제목 '구독여부'(TEXT-C69) 정상. "
     "다만 기본 레이아웃이 이 컬럼을 숨긴다 → FAIL내역 D-01 참조"),
    ("TC-20", "화면", "ALV TEXT 길이", "구독무상(N년이후)/(N년이내)가 온전히 표시",
     "Minor", P, "OUTPUTLEN 22 (C01 985), 실제 값 최대 16자 → 잘림 없음"),
    ("TC-21", "화면", "정렬·필터·엑셀 다운로드", "정렬·필터 정상, 다운로드 파일에도 동일 TEXT",
     "Minor", N,
     "※부분 확인 - ALV 표준 정렬 대화상자의 열 목록(58개)에 '구독여부'가 포함됨을 확인. 정렬 적용·필터·엑셀 다운로드는 미실행"
     "(ALV 열 선택이 스크립팅으로 되지 않고, 파일 다운로드는 별도 승인 필요). 이번 수정과 무관한 ALV 표준기능"),
    ("TC-22", "다품목", "동일 상담번호 복수 품목",
     "품목별 값이 같으면 중복 없이 1행, 다르면 임의 값으로 처리하지 않고 결함으로 기록",
     "Critical", F,
     "동일값(2026091990008, 품목1·2 모두 S/12) → 구독 화면에 1행만, 정상. "
     "혼재(2026091990007, 품목1=S/12, 품목2=공란) → ZCRM177A 에 '무상', ZCRM177AS 에 '구독무상(N년이후)' 로 양쪽 모두 표시. FAIL내역 D-02 참조"),
    ("TC-23", "개인정보", "개인정보 복호화 사용", "권한·옵션에 따라 기존처럼 복호화, SELECT/INTO 순서 밀림 없음",
     "Major", N,
     "※해당없음 - ZCRMR177_1 의 SELECT 목록과 ALV 필드카탈로그에 CITY1·TELF1·TELF2·ENGTEL 이 없다(암호화 대상 개인정보를 다루지 않음). 복호화 옵션 필드도 선택화면에 없음"),
    ("TC-24", "개인정보", "개인정보 비복호화 사용", "기존 비복호화 결과와 동일", "Major", N, "※해당없음 - TC-23 과 같은 사유"),

    ("TC-25", "문서처리", "ZCRM177AS 판매오더 생성", "판매오더 1회 생성, 기존 필드 정상, 구독 TEXT 가 처리 판정에 영향 없음",
     "Critical", N,
     "※미실행(BLOCKED) - ALV 선택 체크박스(CHK)를 SAP GUI 스크립팅으로 켜도 프로그램 내부테이블에 전달되지 않아 항상 E '선택한 데이터가 없습니다'. "
     "modifyCell/modifyCheckbox + Enter/열 이동/행 선택 5가지 조합 모두 동일. 화면에는 체크가 보이나 WORKING_CHECK 의 LOOP AT gt_data WHERE chk='X' 가 0건. "
     "소스상 문서처리 경로의 이번 변경은 GT_AUTO-PROGRAM_NM 3곳(F01 911~918, 1156~1163, 1290~1297)뿐이고 오더/납품/출고 로직은 변경되지 않았다"),
    ("TC-26", "문서처리", "ZCRM177AS 납품문서 생성", "납품문서 생성·연결, 재실행 시 중복 없음", "Critical", N, "※미실행(BLOCKED) - TC-25 와 같은 사유"),
    ("TC-27", "문서처리", "ZCRM177AS 출고전기", "출고전기 완료, 상태·문서번호 갱신, 재고·문서흐름 일치", "Critical", N, "※미실행(BLOCKED) - TC-25 와 같은 사유"),
    ("TC-28", "문서처리", "구독 건 전체 매출처리 연속 실행", "주문·납품·출고가 순서대로 1회씩", "Critical", N, "※미실행(BLOCKED) - TC-25 와 같은 사유"),
    ("TC-29", "문서처리", "구독 복수행 일괄처리", "선택 건만 처리, 상담번호별 문서번호 정확", "Major", N, "※미실행(BLOCKED) - TC-25 와 같은 사유"),
    ("TC-30", "문서처리", "일반 ZCRM177A 문서처리 회귀", "기존 방식대로 문서 생성, 신규 필드로 인한 오류 없음", "Critical", N, "※미실행(BLOCKED) - TC-25 와 같은 사유"),
    ("TC-31", "문서처리", "이미 처리된 건 재처리 방지", "통제 메시지, 추가 문서 미생성", "Major", N, "※미실행(BLOCKED) - TC-25 와 같은 사유"),
    ("TC-32", "문서처리", "처리 중 오류 발생", "오류 메시지, 부분 처리·잘못된 완료표시 없음", "Major", N, "※미실행(BLOCKED) - TC-25 와 같은 사유"),
    ("TC-33", "문서처리", "화면 재조회 정합성", "생성 문서·상태 유지, TEXT 동일, 일반 화면에 구독건 미표시", "Major", N,
     "※문서처리 부분 미실행(BLOCKED). 재조회 시 구독 건이 일반 화면에 나타나지 않는 것은 TC-04·TC-05 에서 반복 확인"),

    ("TC-34", "예외", "WRTDATE 공란", "Native SQL 오류·덤프 없음, SUBSWRTMONTH 기준 TEXT 표시",
     "Major", P, "2026091990006 (S, SUBSWRTMONTH=3, WRTDATE 공란) → 덤프 없이 '구독무상(N년이후)'. WRTDATE 는 애초에 이 프로그램이 읽지 않음"),
    ("TC-35", "예외", "외부 DB 연결 오류", "CX_SY_NATIVE_SQL_ERROR 처리, 불완전 데이터 미저장",
     "Major", N,
     "※미실행 - 연결 오류 모의는 개발 담당 환경이 필요. ZCRMR177_1 자체에는 EXEC SQL 이 없고(전체 소스 검색 0건), "
     "Native SQL 은 적재 FM ZCRMF02_RFC_ZCRM_CICSVC_ALL(기능그룹 ZCRMG02, LZCRMG02U27/U73)에 있으며 이 화면에서는 선택화면의 "
     "'상담번호 동기화(CIC->SAP)' 버튼(F01 3065)으로 호출된다. 검증 대상은 그 FM"),
    ("TC-36", "성능", "대량 조회 성능", "TIME_OUT·메모리 오류 없음, 건수가 원천 분리 기준과 일치",
     "Major", P,
     "조치일자 2026.01.01~09.19, 완료건 포함: ZCRM177A 21,575건 / 약 73초, ZCRM177AS 13건 / 약 19초. TIME_OUT·덤프 없음. "
     "※기존 화면(ZCRM177A)의 대량 조회 시간은 업무 허용범위를 운영과 비교해 확인 권장"),
    ("TC-37", "권한", "권한 검증", "권한 보유자 정상, 미보유자 통제, 기존 권한 영향 없음",
     "Major", N, "※미실행 - 사용자 권한 부여/회수는 Basis 영역이라 테스트 계정으로 수행하지 않음. 테스트 사용자(KKR003-1)는 두 T-code 모두 정상 실행"),
    ("TC-38", "이력", "프로그램명 이력값", "설계에 맞는 PROGRAM_NM 저장",
     "Major", N,
     "※실측 미실행(TC-25 종속). 소스상 ZCRM177AS 로 실행하면 'ZCRM177AS', 그 외에는 'ZCRM177A' 로 저장된다(F01 913~917, 1158~1162, 1292~1296). "
     "기존 로직 '*gt_auto-program_nm = ZCRM177A' 는 주석 처리됨. 분리 저장이 맞는지 업무 확정 필요"),
    ("TC-39", "이력", "내부오더 기준정보 회귀", "확정된 내부오더 적용, 기준정보 미조회 오류 없음",
     "Major", N, "※미실행(TC-25 종속). 소스상 MAKE_ORDER 의 기준정보/계정지정 로직은 이번 개선에서 변경되지 않음"),
    ("TC-40", "화면", "화면 제목 구분", "신규 화면 제목으로 구독 화면임을 확인, 일반과 혼동 없음",
     "Major", P,
     "ZCRM177AS 결과화면 제목 '[구독] AS 비용처리', ZCRM177A 는 'AS 비용처리'. 소스: SET TITLEBAR 'ZSUB' / '0100' 분기(O01 63~70)"),
    ("TC-41", "정합성", "전체 데이터 분할 합계",
     "두 화면 간 중복 없음, 일반 + 구독 = 대상 전체",
     "Critical", F,
     "조치일자 2026.09.19 기준 대상 상담번호 10건(2026091990001~008·012·013)인데 ZCRM177A 4건 + ZCRM177AS 7건 = 11건. "
     "혼재 상담번호 2026091990007 이 양쪽에 모두 나타나 1건 초과. FAIL내역 D-02 참조"),
    ("TC-42", "정합성", "NULL 과 공백 구분", "NULL·공백 모두 일반으로 처리되고 구독 화면에 나타나지 않음",
     "Major", P,
     "SAP 테이블 컬럼이라 NULL 은 존재할 수 없고, 공백 건(001·002·007-2·013)은 모두 ZCRM177A 에서만 조회됨. "
     "'S 외 값'(예 'X')은 SUBS_TYPE 도메인이 거부해 SE16N 저장 자체가 되지 않음 → 현재 값 영역은 공백/S 뿐"),
    ("TC-43", "TEXT", "텍스트 표기 문구 검증",
     "유상 / 무상 / 구독무상(N년이후) / 구독무상(N년이내), 맞춤법·괄호·띄어쓰기 일치",
     "Major", P,
     "네 가지 모두 화면에서 확인(유상·무상은 ZCRM177A, 구독무상 2종은 ZCRM177AS). 소스 리터럴과 화면 표기 동일, 괄호·띄어쓰기 요구사항과 일치"),
    ("TC-44", "정리", "테스트 데이터 및 문서 정리",
     "생성 문서 취소·데이터 원복 기록",
     "Major", P,
     "판매오더·납품·자재문서는 생성되지 않았으므로(TC-25 BLOCKED) 취소 대상 없음. 기존 데이터의 SUBS_TYPE 등을 임의 변경한 건 없음(전부 신규 삽입). "
     "테스트 데이터는 사용자 방침에 따라 KDQ 에 남겨 둠 - 상담번호 2026091990001~013, 생성자 ZTEST0919"),

    ("X-01", "추가", "기본 레이아웃에서의 구독여부 노출",
     "신규 컬럼이 기본 상태에서 사용자에게 보여야 함",
     "Critical", F,
     "두 T-code 모두 전역 기본 레이아웃 '/PT'(기본 설정)이 적용되어 구독여부가 '숨긴 열'에 있다. 레이아웃 변경으로 표시열에 옮겨야만 보인다. FAIL내역 D-01"),
    ("X-02", "추가", "두 T-code 의 ALV 레이아웃 키 공유",
     "화면별로 레이아웃이 독립적이어야 함",
     "Major", F,
     "GV_VARIANT-REPORT = SY-REPID 만 사용(Z00CA_ALV_M 654)하고 HANDLE 을 쓰지 않아 ZCRM177A 와 ZCRM177AS 가 같은 레이아웃 키를 공유한다. "
     "한쪽에서 저장한 기본 레이아웃이 다른 쪽에도 그대로 적용됨. FAIL내역 D-01"),
    ("X-03", "추가", "무상부품 없는 상담번호의 조회 제외",
     "무상 대상이 아닌 건은 조회되지 않아야 함",
     "Minor", P,
     "ZCRM_CICCOST_SUB 에 COST=0 행이 없으면 상담번호가 조회에서 빠진다(F01 165~168, '올유상인건 지운다'). 기존 로직이며 구독/일반 모두 동일하게 동작"),
]

prog_rows = [{"id": c[0], "case": "[%s] %s" % (c[1], c[2]), "expect": c[3], "sev": c[4], "result": c[5], "data": c[6]} for c in CASES]

# ---------------------------------------------------------------- 공통테스트케이스
common = {}


def setc(nums, target, result, note=""):
    for n in nums:
        common["UC-%03d" % n] = (target, result, note)


setc(range(1, 120), "N", "", "해당 없음")
setc([1], "Y", P, "전기일자·회사코드·판매조직 필수(OBLIGATORY), 기본값 오늘/1100/ASKR")
setc([8], "Y", P, "조치일자 기본값(당월), 완료건 포함 체크박스 기본 해제")
setc([9, 10, 11], "Y", P, "조회·레이아웃·정렬 대화상자 전 구간 덤프·행업 없음")
setc([16], "Y", P, "동일 상담번호는 상담번호 단위로 1행(DELETE ADJACENT DUPLICATES)")
setc([40], "Y", P, "조회조건 필수값 체크 정상")
setc([43], "Y", F, "일반 4건 + 구독 7건 = 11건 > 대상 10건(혼재 상담번호 중복, TC-41)")
setc([44], "Y", P, "ZCRM177AS=구독만, ZCRM177A=일반만(TC-04·05)")
setc([54], "Y", P, "T-code 에 따라 조회 대상만 바뀌고 선택화면 구성은 동일")
setc([55], "Y", P, "조치일자 + 구독구분 조건으로 분리")
setc([57], "Y", P, "상담번호당 1행")

# ---------------------------------------------------------------- FAIL 내역
fail_rows = [
    {"id": "D-01", "case": "신규 '구독여부' 컬럼이 기본 레이아웃에서 숨겨져 있음 (TC-19 / X-01 / X-02)",
     "tester": "KKR003-1",
     "detail": "ZCRM177A·ZCRM177AS 를 실행하면 전역 기본 레이아웃 '/PT'(레이아웃 선택 팝업에서 '기본 설정')가 적용되고, 이번에 추가한 '구독여부' 컬럼은 '숨긴 열' 목록에 들어간다. "
               "사용자가 레이아웃 변경으로 직접 옮기지 않으면 구독여부를 볼 수 없어 개선 목적(F. ALV 구독여부 TEXT 표시)이 기본 상태에서 달성되지 않는다.\n"
               "원인: 필드카탈로그에는 SUBS_TEXT 가 정상 추가되지만(ZCRMR177_1C01 979~987), ALV 가 IS_VARIANT = GV_VARIANT 로 저장된 레이아웃을 적용하고 "
               "GV_VARIANT-REPORT 는 SY-REPID 만 세팅한다(Z00CA_ALV_M 654). 그 결과 ① 기존에 저장된 기본 레이아웃이 신규 컬럼을 숨기고 ② 두 T-code 가 같은 레이아웃 키를 공유한다.\n"
               "조치안: (1) 기본 레이아웃 '/PT' 에 구독여부를 추가해 다시 저장하거나, (2) 화면별로 GV_VARIANT-HANDLE 을 분리(예: 구독 모드일 때 'SUBS')해 레이아웃을 독립시킨다.",
     "images": ["S20_r177_lay_layout.png", "S20_r177_lay2_col0.png"]},
    {"id": "D-02", "case": "다품목 상담번호에 구독/비구독이 섞이면 양쪽 화면에 중복 표시되고 구독여부가 화면마다 다름 (TC-22 / TC-41)",
     "tester": "KKR003-1",
     "detail": "상담번호 2026091990007 은 품목1 SUBS_TYPE='S'(SUBSWRTMONTH 12), 품목2 공란이다. ZCRM177A 에서는 '무상', ZCRM177AS 에서는 '구독무상(N년이후)' 으로 "
               "같은 상담번호가 양쪽 화면에 모두 나타난다. 그 결과 조치일자 2026.09.19 기준 대상 10건에 대해 4건 + 7건 = 11건이 되어 분할 합계가 맞지 않고, "
               "같은 건이 두 화면에서 각각 매출처리 대상으로 잡힐 수 있다.\n"
               "원인: 조회는 품목(ZCRM_CICSVC) 단위로 SUBS_TYPE 을 걸러내고(ZCRMR177_1F01 116), 표시는 상담번호 단위로 중복을 제거한다"
               "(SORT gt_data BY csltno / DELETE ADJACENT DUPLICATES COMPARING csltno, F01 143~144). 남는 행이 무엇인지에 따라 구독여부가 달라진다.\n"
               "조치안: 상담번호 단위 대표값 기준을 먼저 확정하고(예: 품목 중 하나라도 'S' 면 구독 건으로 본다) 그 기준으로 상담번호 단위 분기를 적용한다. "
               "기준이 정해지기 전까지는 혼재 건을 오류로 표시하거나 별도 목록으로 분리하는 방법도 가능하다.",
     "images": ["S20_TC04_A2_tc04_alv.png", "S20_TC05_S_tc05_alv.png"]},
]

# ---------------------------------------------------------------- 증적
ev = [
    ("P-01", "SE16N - 구독 테스트 데이터 삽입(ZCRM_CICSVC 2026091990003)", "S19_r177_svc_ZCRM_CICSVC_2026091990003__saved.png"),
    ("P-01", "원천값 ZCRM_CICSVC - SUBS_TYPE / SUBSWRTMONTH / WRTDATE (13행)", "S12_ZCRM_CICSVC_r177_src.png"),
    ("P-01", "원천값 ZCRM_CICCOST_M - 조치일자 / 유무상(ISFREE) (12행)", "S12_ZCRM_CICCOST_M_r177_srcm.png"),
    ("TC-03", "ZCRM177A 선택화면(기존)", "S20_TC03_sel_sel_177a.png"),
    ("TC-02", "ZCRM177AS 선택화면(신규)", "S20_TC02_sel_sel_177as.png"),
    ("TC-04", "ZCRM177A 결과 - 일반 4건, 구독여부 유상/무상", "S20_TC04_A2_tc04_alv.png"),
    ("TC-05", "ZCRM177AS 결과 - 구독 7건, 제목 [구독] AS 비용처리", "S20_TC05_S_tc05_alv.png"),
    ("TC-06", "조치일자 2026.09.17~19 - 시작일 건 포함 / 09.16 건 제외", "S20_TC06_S_tc06_alv.png"),
    ("TC-10", "SUBSWRTMONTH 는 INT2(부호 있음) - SE11 필드 정의", "S23_r177_type4_46.png"),
    ("TC-10", "SE16N 음수 입력 시도 - 저장되지 않음", "S21_r177_neg_m6_after.png"),
    ("TC-10", "SM30 - ZCRM_CICSVC 유지보수 다이얼로그 없음", "S20_r177_sm30c_status_5.png"),
    ("TC-19", "레이아웃 선택 - 전역 레이아웃 /PT 가 기본 설정", "S20_r177_lay_layout.png"),
    ("TC-19", "레이아웃 변경 - '구독여부'가 숨긴 열에 있음(내부코드 SUBS_TYPE 은 목록에 없음)", "S20_r177_lay2_col0.png"),
    ("TC-21", "ALV 정렬 대화상자 열 목록에 '구독여부' 포함", "S20_TC21_sortdlg_tc21_sortdlg.png"),
    ("TC-25", "체크박스는 화면에 표시되지만", "S20_TC25_tc25_checked.png"),
    ("TC-25", "주문생성 시 E '선택한 데이터가 없습니다' - 스크립팅으로 수행 불가", "S20_TC25_tc25_after.png"),
    ("TC-36", "ZCRM177AS 2026.01.01~09.19 대량 조회 - 덤프 없음", "S20_TC36_perf_S_tc36_s.png"),
    ("TC-41", "ZCRM177A 2026.09.19 건수(4건)", "S20_TC41A5_count_a.png"),
]
ev_out = []
for tc, cap, fn in ev:
    t = thumb(fn)
    if t:
        ev_out.append({"tc": tc, "caption": cap, "path": t})
for fr in fail_rows:
    fr["image_paths"] = [p for p in (thumb(f, 620) for f in fr["images"]) if p]

payload = {
    "date": DATE,
    "cover": {"C17": "2026.09.19"},
    "revision": {"B5": "0.01", "C5": "ZCRMR177_1 일시불구독 개선(ZCRM177A / ZCRM177AS) 일반·구독 조회 분리 테스트 최초 작성 (48건: PASS 25 / FAIL 4 / 미실행 19)", "D5": "2026.09.19",
                 "E5": "※작성자 기입 필요", "F5": "", "G5": "",
                 "B6": "0.02", "C6": "적재 FM(ZCRMF02_RFC_ZCRM_CICSVC_ALL / 기능그룹 ZCRMG02) 소스 확인 후 판정 정정 - WRTDATE 연동은 요구사항대로 구현되어 있어 D-03 철회, TC-08 PASS, TC-12·13·35 근거 보강 (48건: PASS 26 / FAIL 4 / 미실행 18, FAIL내역 3건→2건)",
                 "D6": "2026.09.19", "E6": "※작성자 기입 필요", "F6": "", "G6": ""},
    "prog_header": {"B1": "서비스 - AS 무상 매출처리", "D1": "※개발자명 기입 필요",
                    "B2": "SR03000035400 ZCRMR177_1 일시불구독 개선 - 일반/구독 조회 분리 및 구독여부 표시 (ZCRM177A / ZCRM177AS)", "D2": "ZCRMR177_1"},
    "prog_rows": prog_rows,
    "common": {k: {"target": v[0], "result": v[1], "note": v[2]} for k, v in common.items()},
    "fail_rows": fail_rows,
    "evidence": ev_out,
    "evidence_sheet": "테스트증적(캡처)",
    "legend": {"blank": "결과란이 비어 있으면 미실행이며 사유는 '테스트 데이터/실제결과' 열에 적었다(※해당없음 / ※미실행(BLOCKED)). "
                        "X-xx 는 시나리오 외 추가 테스트다. 케이스별 화면은 '테스트증적(캡처)' 시트, 따라하기용 데이터는 '테스트데이터' 시트에 있다.",
               "note": "KDQ 210, 2026-09-19, 사용자 KKR003-1. 수정 전/후 구분은 소스의 '**<<SR03000035400 … / **>>SR03000035400 …' 주석 블록 13곳을 기준으로 했고 "
                       "전체 소스는 RS_ABAP_SOURCE_SCAN 으로 내려받아 results\\src\\ZCRMR177_1_diffblocks.txt 에 정리했다. "
                       "문서처리 계열(TC-25~33, 38, 39)은 ALV 선택 체크박스를 스크립팅으로 프로그램에 전달할 수 없어 미실행이며, 수동 테스트가 필요하다. "
                       "연동(TC-07~13, 35)은 화면 프로그램이 아니라 적재 FM ZCRMF02_RFC_ZCRM_CICSVC_ALL(기능그룹 ZCRMG02)의 Native SQL 소관이며, "
                       "해당 소스도 같은 주석 표식으로 확인해 results\\src\\ZCRMG02_subs_blocks.txt 에 정리했다."},
}
io.open(OUT, "w", encoding="utf-8").write(json.dumps(payload, ensure_ascii=False, indent=1))
res = [r["result"] for r in prog_rows]
print("prog rows:", len(res), "PASS", res.count("PASS"), "FAIL", res.count("FAIL"), "미실행", res.count(""))
cr = list(payload["common"].values())
print("common Y:", sum(1 for v in cr if v["target"] == "Y"), "PASS", sum(1 for v in cr if v["result"] == "PASS"), "FAIL", sum(1 for v in cr if v["result"] == "FAIL"))
print("evidence:", len(ev_out), "missing:", missing)
