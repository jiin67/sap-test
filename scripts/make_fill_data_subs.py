# make_fill_data_subs.py : JSON payload for the company unit-test template
# 2026-09-11 AS 무상지원금 일시불구독 개선 (ZCRMR058_1 / ZCRMR060_1 / RFC ZCRMF03_RFC_ZCRM_CHRG_CNFRM)
import io, json, os
from PIL import Image as PILImage

OUT = r"C:\sap-test\results\_fill_data_subs.json"
SHOTS = r"C:\sap-test\results\shots"
THUMB = r"C:\sap-test\results\shots\_tpl_subs"
os.makedirs(THUMB, exist_ok=True)

DATE = "2026-09-11"


def thumb(fn, maxw=760):
    src = os.path.join(SHOTS, fn)
    if not os.path.exists(src):
        print("  missing shot:", fn)
        return None
    dst = os.path.join(THUMB, fn)
    im = PILImage.open(src)
    im.thumbnail((maxw, maxw))
    im.save(dst, optimize=True)
    return dst


prog_header = {
    "B1": "서비스 - AS 무상지원금",
    "D1": "※개발자명 기입 필요",
    "B2": "AS 무상지원금 관리 - 일시불구독 개선 (ZCRM058S / 059S / 060S, 확정일 RFC)",
    "D2": "ZCRMR058_1 / ZCRMR060_1 / ZCRMF03_RFC_ZCRM_CHRG_CNFRM",
}

DATA = ("기간월 2026.07 / X=117786(CT11) 구독 3건 2026070100278·0450·0862, "
        "Y=100011(CT02) 구독 1건 2026070102721")
RFCBUG = "→ FAIL내역 UT-SUB-006 참조 (RFC 가 받은 CSLTNO 가 3자리 밀림)"

# id, 테스트 케이스, 예상결과, 심각도, 테스트 데이터/비고, 결과
prog_rows = [
    ("UT-SUB-001", "[1-1] 구독 처리화면 진입 (ZCRM058S)",
     "SUBS_TYPE='S' 인 상담건만 집계되어 표시된다", "Critical",
     DATA + ". 세팅 전 0행 → 세팅 후 X 1행 20,000원(구독 3건). 같은 대리점 일반건 190,000원은 미포함",
     "PASS"),
    ("UT-SUB-002", "[1-2] 일반 화면 영향 없음 (ZCRM058A)",
     "비구독 건만 표시되고 구독 건은 보이지 않는다", "Critical",
     "X 190,000 → 170,000 (−20,000), Y 65,000 → 60,000 (−5,000) = 구독분만 빠짐. "
     "058A + 058S 합계 = 변경 전 058A", "PASS"),
    ("UT-SUB-003", "[1-3] 구독 조회전용 (ZCRM059S)",
     "구독 건이 표시되고 처리 버튼은 표시되지 않는다", "Major",
     "X 1행 20,000. 툴바는 지급대상데이터조회·상세리스트조회·회계전표출력만 표시. "
     "무상지원금확정/취소, 대리점 확정/확정취소, 일괄회계전표생성/삭제 없음", "PASS"),
    ("UT-SUB-004", "[1-4] 구독 상세 직접 조회 (ZCRM060S)",
     "구독 상담건의 상세만 표시된다", "Major",
     "X: 기사 3명(0000130784/0000132094/0000134384 = 구독 3건의 기사), 합계 20,000. "
     "060A 는 기사 8명 합계 170,000", "PASS"),
    ("UT-SUB-005", "[1-5] 상세 드릴다운 (ZCRM058S → 상세리스트조회)",
     "구독 상세(ZCRMR060_1, 구독모드)로 이동한다", "Major",
     "058S X 행 선택 후 상세리스트조회 → ZCRMR060_1 7행, 합계 20,000 (= ZCRM060S 결과)", "PASS"),
    ("UT-SUB-006", "[2-1] RFC 정상 확정등록",
     "IV_ACTION=S, 유효 상담번호 3건 → EV_RESULT='S', 3건 저장", "Critical",
     "입력 2026070100862/2026070102721/2026070100278. EV_RESULT=E, 3건 모두 '존재하지 않는 상담번호', "
     "반환 CSLTNO 6070100862202 등. 저장 0건 " + RFCBUG, "FAIL"),
    ("UT-SUB-007", "[2-2] RFC 확정취소",
     "IV_ACTION=C, 2건 → EV_RESULT='S', 2건 삭제", "Critical",
     "2026070100278/2026070100450 (SE16N 으로 넣어둔 확정 행). EV_RESULT=E, 반환 CSLTNO "
     "6070100278000/6070100450000 '존재하지 않는 상담번호'. 삭제 0건 " + RFCBUG, "FAIL"),
    ("UT-SUB-008", "[2-3] RFC 처리구분 누락",
     "IV_ACTION 공백 → 'E', 처리구분 미입력", "Major",
     "※미실행 - SE37 테스트 프레임이 필수 매개변수 IV_ACTION 공백 실행을 막음('매개변수값 입력' 팝업). "
     "RFC 내부 검증은 레거시 호출로 확인 필요", ""),
    ("UT-SUB-009", "[2-4] RFC 잘못된 처리구분",
     "IV_ACTION='X' → 'E', 유효하지 않은 구분", "Major",
     "EV_RESULT=E, ET_RETURN '유효하지 않은 처리구분입니다 (S 또는 C)'", "PASS"),
    ("UT-SUB-010", "[2-5] RFC 없는 상담번호 포함 (부분성공)",
     "없는 상담번호만 'E', 나머지는 저장되고 EV_RESULT='E'", "Critical",
     "2026070100278 + 9999999999999. 유효건까지 '존재하지 않는 상담번호'(6070100278202), 저장 0건 " + RFCBUG,
     "FAIL"),
    ("UT-SUB-011", "[2-6] RFC 확정일 누락",
     "IV_ACTION=S 인데 CFDATE1 공백 → 해당건 'E'(확정일자 누락)", "Major",
     "2026070100450 / 공백. 'E' 이나 사유가 '존재하지 않는 상담번호'(6070100450000) " + RFCBUG, "FAIL"),
    ("UT-SUB-012", "[2-7] RFC 미래일자",
     "CFDATE1 미래일자 → 해당건 'E'(미래일자 불가)", "Major",
     "2026070100450 / 2026.12.31. 'E' 이나 사유가 '존재하지 않는 상담번호'(6070100450202) " + RFCBUG, "FAIL"),
    ("UT-SUB-013", "[2-8] RFC 빈 데이터",
     "IT_DATA 없음 → 'E', 전송 데이터 없음", "Minor",
     "EV_RESULT=E, ET_RETURN '전송 데이터가 없습니다'", "PASS"),
    ("UT-SUB-014", "[2-9] RFC 중복 상담번호",
     "같은 CSLTNO 2줄 → 중복 제거되어 1건 저장", "Major",
     "2026070100862 2줄. 2줄 모두 '존재하지 않는 상담번호'(6070100862202), 저장 0건 " + RFCBUG, "FAIL"),
    ("UT-SUB-015", "[3-1] 전건확정 - 미확정 (0/3)",
     "ZCRM058S 대리점확정일(CFDATE1) 공백", "Critical",
     "ZCRM_CHRG_CNFRM 0건 상태에서 058S X CFDATE1 공백", "PASS"),
    ("UT-SUB-016", "[3-2] 전건확정 - 부분확정 (2/3)",
     "3건 중 2건만 확정 → CFDATE1 여전히 공백", "Critical",
     "※RFC 결함으로 확정 데이터를 SE16N 으로 직접 입력(ERNAM=ZTEST0911, 사용자 승인). "
     "X-1 2026.08.05, X-2 2026.08.06 → 058S X CFDATE1 공백", "PASS"),
    ("UT-SUB-017", "[3-3] 전건확정 - 전건확정 (3/3)",
     "3건 모두 확정 → CFDATE1 = 확정일 최대값", "Critical",
     "X-3 2026.08.07 추가 → 058S X CFDATE1 2026.08.07 (최대값). 같은 시점 058A X CFDATE1 은 공백 유지", "PASS"),
    ("UT-SUB-018", "[3-4] 전건확정 - 확정 후 취소",
     "전건확정에서 1건 취소 → CFDATE1 다시 공백", "Critical",
     "X-3 확정 행 삭제(3/3 → 2/3) → 058S X CFDATE1 공백", "PASS"),
    ("UT-SUB-019", "[3-5] 전건확정 - 단일건 대리점",
     "대상 1건인 대리점은 1건 확정 즉시 CFDATE1 표시", "Major",
     "Y-1 2026.08.04 (SPMON 2026.07) → 058S Y CFDATE1 2026.08.04", "PASS"),
    ("UT-SUB-020", "[4-1] 구독 수동확정 차단 (ZCRM058S)",
     "대리점 확정/확정취소(KUNCR/KUNCC) 버튼이 없다", "Major",
     "058S 툴바에 대리점 확정(F2)·대리점 확정취소(F5) 없음. 기능코드는 GUI 스크립팅으로 직접 확인 불가해 "
     "버튼·단축키 부재로 판정", "PASS"),
    ("UT-SUB-021", "[4-2] 일반 수동확정 유지 (ZCRM058A)",
     "대리점 확정/확정취소 버튼이 정상 표시된다", "Major",
     "058A 툴바에 대리점 확정(F2)·대리점 확정취소(F5) 표시·활성", "PASS"),
    ("UT-SUB-022", "[4-3] 구독 전표생성",
     "전건확정 후 일괄회계전표생성 → 전표 정상 생성", "Critical",
     "※미실행 - 실제 회계전표가 생성되고, ZCRM_ASSPTFEEH1 키(SPMON+KUNNR)를 같은 대리점의 일반건과 "
     "공유하는 점을 개발자 확인 후 진행하기로 함(사용자 결정)", ""),
    ("UT-SUB-023", "[4-4] 미확정 전표 차단",
     "부분확정 상태에서 일괄회계전표생성 → 확정일자 없음 오류", "Critical",
     "※미실행 - 버튼 클릭 시 '계속 진행 하시겠습니까?' 확인 팝업이 확정일자 검사보다 먼저 표시됨. "
     "[예] 선택 시 게이트가 동작하지 않으면 실전표가 생기므로 취소(사용자 결정). "
     "취소 후 전표·ZCRM_ASSPTFEEH1 변화 없음 확인", ""),
    ("UT-SUB-024", "[5-1] RFC RSLTDATE 공백 상담건",
     "기간월 산출 불가 오류", "Major",
     "※미실행 - RFC 결함(UT-SUB-006)으로 상담번호 검증 단계에서 모두 거부되어 검증 불가", ""),
    ("UT-SUB-025", "[5-2] SPMON 경계",
     "확정 행의 SPMON 이 기간월(pa_spmon)과 같을 때만 표시된다", "Major",
     "Y-1 을 SPMON 2026.06 으로 입력 → 058S(2026.07) CFDATE1 공백, 2026.07 로 변경 → 2026.08.04 표시. "
     "※RFC 의 RSLTDATE 기반 SPMON 파생은 결함으로 미검증", "PASS"),
    ("UT-SUB-026", "[5-3] 테스트 후 원복",
     "SUBS_TYPE 원복 후 058A 정상 복귀, 구독 화면에 미표시", "Critical",
     "SUBS_TYPE 4건 공백 복원(전체 'S' 0건), ZCRM_CHRG_CNFRM 0건. 058A X 190,000 / Y 65,000 (변경 전과 동일), "
     "058S X·Y 0행, 060S X 0행", "PASS"),
    ("UT-SUB-027", "[회귀] 일반 화면에 구독 확정일 미반영",
     "구독 전건확정 상태에서도 ZCRM058A 의 CFDATE1 은 바뀌지 않는다", "Major",
     "3-3 상태(X 3/3 확정)에서 058A X CFDATE1 공백, 3-5 상태에서 058A Y CFDATE1 공백", "PASS"),
]

# ---------------------------------------------------------------- 공통테스트케이스
P, F = "PASS", "FAIL"
common = {}


def setc(nums, target, result, note=""):
    for n in nums:
        common["UC-%03d" % n] = (target, result, note)


NA = "해당 없음"
# 입력 (058 선택화면 / RFC 입력)
setc([1], "Y", P, "058 기간월·센터 미입력 시 '모든 필수 입력 필드에 값을 입력하십시오.' 로 실행 차단")
setc([2], "N", "", NA)
setc([3], "Y", F, "RFC 처리구분·빈 데이터 검증은 정상이나 유효 상담번호까지 거부 → UT-SUB-006")
setc([4], "N", "", NA)
setc([5], "Y", "", "미실행 - RFC 미래일자 검증(2-7)은 결함으로 미도달")
setc([6, 7], "N", "", NA)
setc([8], "Y", P, "라디오 기본값(058 '전체', 060 '조회') 확인")
setc([9, 10, 11], "Y", P, "전 실행 구간(058/059/060, SE37 RFC)에서 덤프·행업 없음")
setc([12], "N", "", NA)
setc([13], "Y", F, "RFC 등록(S) 요청이 한 건도 저장되지 않음 → UT-SUB-006")
setc([14], "Y", "", "미실행 - RFC 저장 불가로 저장값 대조 불가")
setc([15], "N", "", NA)
setc([16], "Y", "", "미실행 - 중복 상담번호(2-9)는 결함으로 중복 제거 로직 미도달")
setc(range(17, 24), "N", "", NA)
# 수정
setc(range(24, 30), "N", "", NA)
setc([30], "Y", "", "미실행 - RFC 재등록(변경) 결함으로 미검증")
setc(range(31, 40), "N", "", NA)
# 조회
setc([40], "Y", P, "필수 조회조건(기간월·센터) 체크 확인")
setc([41, 42], "N", "", NA)
setc([43], "Y", P, "058S 구독 3건 20,000 = 060S 기사 3명 합계 20,000")
setc([44], "Y", P, "058A + 058S = 변경 전 058A (190,000 = 170,000 + 20,000 / 65,000 = 60,000 + 5,000)")
setc([45], "Y", P, "확정 데이터 입력·변경·삭제가 058S CFDATE1 에 즉시 반영 (SE16N 입력 기준)")
setc([46], "Y", P, "058S → 상세리스트조회 드릴다운 정상")
setc([47, 48], "N", "", NA)
setc([49, 50, 51], "Y", P, "")
setc([52], "N", "", NA)
setc([53], "Y", P, "구독/일반 분리 조회 합계가 변경 전 전체와 일치")
setc([54], "Y", P, "058S 에서 넘긴 구독모드가 060 상세 조회에 연동")
setc([55], "Y", P, "기간월 + 센터 + 대리점 다중조건 적용 확인")
setc([56], "N", "", NA)
setc([57], "Y", P, "대리점당 1행, 중복 표시 없음")
# 삭제 (RFC 취소)
setc([58, 59], "N", "", NA)
setc([60], "Y", P, "059S 조회전용 처리버튼 미표시, 058S 대리점 수동확정 버튼 미표시")
setc(range(61, 66), "N", "", NA)
setc([66], "Y", F, "RFC 취소(C) 2건이 삭제되지 않음 → UT-SUB-006")
setc(range(67, 70), "N", "", NA)
# 출력
setc(range(70, 87), "N", "", "해당 없음 - 출력 기능 변경 없음")
# 그래픽
setc([87], "Y", P, "개발용 메시지 노출 없음")
setc([88], "Y", P, "T-code 별 처리버튼 표시 제어 정상 (058S/059S)")
setc([89], "N", "", NA)
setc([90], "Y", P, "확정 데이터 변경 후 재조회 시 CFDATE1 갱신")
setc([91, 92], "N", "", NA)
setc([93], "Y", "", "판정 보류 - 060 결과화면 제목 건수가 행이 있어도 (0). 060A 도 동일해 기존 동작으로 추정, 개발자 확인 필요")
setc([94, 95, 96], "N", "", NA)
setc([97], "Y", P, "일괄회계전표생성 확인 팝업 예/아니오/취소 정상 표시")
setc([98], "N", "", NA)
setc([99], "Y", F, "RFC 가 확정일 누락·미래일자 건에 '존재하지 않는 상담번호' 로 잘못 응답 → UT-SUB-006")
setc([100], "N", "", NA)
setc([101], "Y", P, "라디오 초기값 확인")
setc(range(102, 106), "N", "", NA)
setc([106], "Y", P, "058S 대리점 확정 버튼, 059S 처리버튼 미표시")
setc([107], "Y", "", "미실행 - 숨긴 기능코드 직접 호출은 미검증")
setc(range(108, 111), "N", "", NA)
setc([111], "Y", P, "RFC 반환 메시지 한글 사용자 메시지")
setc([112], "N", "", NA)
# 시스템
setc([113], "N", "", NA)
setc([114], "Y", F, "확정일 수신 RFC 결함으로 레거시(CIC) 확정 연계 불가 → UT-SUB-006")
setc(range(115, 120), "N", "", NA)

assert len(common) == 119, len(common)

# ---------------------------------------------------------------- FAIL내역
BUG = (
    "ZCRMF03_RFC_ZCRM_CHRG_CNFRM 이 유효한 상담번호를 모두 '존재하지 않는 상담번호' 로 거부한다. "
    "등록(S)·취소(C) 모두 한 건도 반영되지 않는다.\n"
    "[현상] ET_RETURN 에 돌아온 CSLTNO 가 입력값과 다르다. 입력 CSLTNO(13자리)+CFDATE1(8자리)를 이어 붙인 "
    "문자열의 4~16번째 문자가 CSLTNO 로 쓰인다. 예) 2026070100278 + 20260805 → 6070100278202, "
    "2026070100450 + 공백(00000000) → 6070100450000. 확인한 9개 입력 모두 같은 규칙이다.\n"
    "[입력 확인] SE37 테스트 화면의 IT_DATA 행은 CSLTNO / CFDATE1 두 필드뿐이며 MANDT 가 없다. "
    "입력은 인터페이스대로 들어갔다.\n"
    "[추정 원인] {CSLTNO, CFDATE1} 행을 MOVE-CORRESPONDING 이 아닌 단순 MOVE 로 ZCRM_CHRG_CNFRM 형 작업영역"
    "(MANDT 3 + CSLTNO 13 + …)에 옮기면 MANDT 가 앞 3자리를 가져가 정확히 이 결과가 된다. 소스 확인 필요.\n"
    "[영향] 2-1, 2-2, 2-5, 2-6, 2-7, 2-9 FAIL, 5-1 검증 불가. 레거시(CIC) 확정일 연계가 동작하지 않는다. "
    "058 의 전건확정 표시 로직(그룹 3)은 SE16N 으로 확정 데이터를 넣어 별도 검증했고 정상이다."
)
fail_rows = [
    {"id": "UT-SUB-006", "case": "[2-1] RFC 정상 확정등록", "tester": "※테스트 수행자 기입 필요",
     "detail": BUG, "images": ["S14_TC2-1_1_it_data.png", "S14_TC2-1_4_et_return.png"]},
    {"id": "UT-SUB-007", "case": "[2-2] RFC 확정취소", "tester": "※테스트 수행자 기입 필요",
     "detail": "UT-SUB-006 과 같은 원인. C 로 2026070100278/0450 취소 요청 → 반환 CSLTNO 6070100278000 / "
               "6070100450000, '존재하지 않는 상담번호'. SE16N 으로 넣어둔 확정 행이 그대로 남음.",
     "images": ["S14_TC2-2_4_et_return.png"]},
    {"id": "UT-SUB-010", "case": "[2-5] RFC 없는 상담번호 포함", "tester": "※테스트 수행자 기입 필요",
     "detail": "UT-SUB-006 과 같은 원인. 유효건 2026070100278 까지 거부되어 부분성공이 확인되지 않음.",
     "images": ["S14_TC2-5_4_et_return.png"]},
    {"id": "UT-SUB-011", "case": "[2-6] RFC 확정일 누락", "tester": "※테스트 수행자 기입 필요",
     "detail": "UT-SUB-006 과 같은 원인. 기대 메시지 '확정일자 누락' 대신 '존재하지 않는 상담번호'.",
     "images": ["S14_TC2-6_4_et_return.png"]},
    {"id": "UT-SUB-012", "case": "[2-7] RFC 미래일자", "tester": "※테스트 수행자 기입 필요",
     "detail": "UT-SUB-006 과 같은 원인. 기대 메시지 '미래일자 불가' 대신 '존재하지 않는 상담번호'.",
     "images": ["S14_TC2-7_4_et_return.png"]},
    {"id": "UT-SUB-014", "case": "[2-9] RFC 중복 상담번호", "tester": "※테스트 수행자 기입 필요",
     "detail": "UT-SUB-006 과 같은 원인. 2줄 모두 거부되어 중복 제거 여부 확인 불가.",
     "images": ["S14_TC2-9_4_et_return.png"]},
]

# ---------------------------------------------------------------- 증적 캡처
evidence = [
    ("UT-SUB-001", "변경 전 ZCRM058S X - 0행 (구독 데이터 없음)", "SUBS_pre_058S_X.png"),
    ("UT-SUB-001", "구독 세팅 후 ZCRM058S X - 1행 20,000원", "SUBS_post_058S_X.png"),
    ("UT-SUB-002", "변경 전 ZCRM058A X - 190,000원, 대리점 확정/확정취소 버튼 표시", "SUBS_pre_058A_X.png"),
    ("UT-SUB-002", "구독 세팅 후 ZCRM058A X - 170,000원", "SUBS_post_058A_X.png"),
    ("UT-SUB-003", "ZCRM059S X - 구독 1행, 처리버튼 없음", "SUBS_post_059S_X.png"),
    ("UT-SUB-004", "ZCRM060S X - 구독 3건의 기사 3명, 합계 20,000원", "SUBS_post_060S_X.png"),
    ("UT-SUB-004", "ZCRM060A X - 일반 기사 8명, 합계 170,000원", "SUBS_post_060A_X.png"),
    ("UT-SUB-005", "ZCRM058S → 상세리스트조회 드릴다운 결과", "SUBS_post_058S_X_drill_drill.png"),
    ("UT-SUB-006", "SE37 IT_DATA 입력 (CSLTNO / CFDATE1 두 필드)", "S14_TC2-1_1_it_data.png"),
    ("UT-SUB-006", "SE37 ET_RETURN - CSLTNO 3자리 밀림, 존재하지 않는 상담번호", "S14_TC2-1_4_et_return.png"),
    ("UT-SUB-007", "RFC 확정취소(C) ET_RETURN", "S14_TC2-2_4_et_return.png"),
    ("UT-SUB-008", "SE37 - IV_ACTION 공백 시 '매개변수값 입력' 팝업", "S14_TC2-3_popup.png"),
    ("UT-SUB-009", "RFC 잘못된 처리구분 ET_RETURN", "S14_TC2-4_4_et_return.png"),
    ("UT-SUB-010", "RFC 없는 상담번호 포함 ET_RETURN", "S14_TC2-5_4_et_return.png"),
    ("UT-SUB-011", "RFC 확정일 누락 ET_RETURN", "S14_TC2-6_4_et_return.png"),
    ("UT-SUB-012", "RFC 미래일자 ET_RETURN", "S14_TC2-7_4_et_return.png"),
    ("UT-SUB-013", "RFC 빈 데이터 ET_RETURN - 전송 데이터가 없습니다", "S14_TC2-8_4_et_return.png"),
    ("UT-SUB-014", "RFC 중복 상담번호 ET_RETURN", "S14_TC2-9_4_et_return.png"),
    ("UT-SUB-016", "ZCRM_CHRG_CNFRM 2/3 입력 (SE16N, ZTEST0911)", "S15_stepA_after_save.png"),
    ("UT-SUB-016", "부분확정 - ZCRM058S X CFDATE1 공백", "SUBS_t32_058S_X.png"),
    ("UT-SUB-017", "ZCRM_CHRG_CNFRM 3/3 + Y(SPMON 2026.06) 입력", "S15_stepB_after_save.png"),
    ("UT-SUB-017", "전건확정 - ZCRM058S X CFDATE1 2026.08.07", "SUBS_t33_058S_X.png"),
    ("UT-SUB-018", "X-3 삭제 후 ZCRM058S X CFDATE1 공백", "SUBS_t34_058S_X.png"),
    ("UT-SUB-019", "단일건 대리점 - ZCRM058S Y CFDATE1 2026.08.04", "SUBS_t35_058S_Y.png"),
    ("UT-SUB-023", "일괄회계전표생성 클릭 시 확인 팝업 (취소함)", "SUBS_t44_058S_X_fipos_popup1.png"),
    ("UT-SUB-025", "SPMON 2026.06 확정 행 - ZCRM058S(2026.07) Y CFDATE1 공백", "SUBS_t52n_058S_Y.png"),
    ("UT-SUB-026", "원복 후 ZCRM058A X - 190,000원", "SUBS_r53_058A_X.png"),
    ("UT-SUB-026", "원복 후 ZCRM058S X - 0행", "SUBS_r53_058S_X.png"),
    ("UT-SUB-027", "전건확정 상태의 ZCRM058A X - CFDATE1 공백", "SUBS_t33_058A_X.png"),
]

ev_out = []
for tc, cap, fn in evidence:
    t = thumb(fn)
    if t:
        ev_out.append({"tc": tc, "caption": cap, "path": t})

for fr in fail_rows:
    fr["image_paths"] = [p for p in (thumb(f, 620) for f in fr["images"]) if p]

payload = {
    "date": DATE,
    "cover": {"C17": "2026.09.11"},
    "revision": {"B5": "0.01", "C5": "최초 작성 - 일시불구독 개선 단위테스트 (KDQ 210)",
                 "D5": "2026.09.11", "E5": "※작성자 기입 필요", "F5": "", "G5": ""},
    "prog_header": prog_header,
    "prog_rows": [
        {"id": a, "case": b, "expect": c, "sev": d, "data": e, "result": f}
        for a, b, c, d, e, f in prog_rows
    ],
    "common": {k: {"target": v[0], "result": v[1], "note": v[2]} for k, v in common.items()},
    "fail_rows": fail_rows,
    "evidence": ev_out,
    "evidence_sheet": "테스트증적(캡처)",
    "legend": {
        "blank": "결과란이 비어 있으면 미실행이다. 사유는 테스트 데이터 열에 기재했다.",
        "note": "대상: ZCRMR058_1 / ZCRMR060_1 / ZCRMF03_RFC_ZCRM_CHRG_CNFRM (KDQ 210, 2026-09-11). "
                "구독 테스트 데이터는 ZCRM_CICSVC_NEW-SUBS_TYPE 을 4건만 'S' 로 바꿔 만들었고(사용자 승인, SE16N), "
                "그룹 3 확정 데이터는 RFC 결함으로 ZCRM_CHRG_CNFRM 에 SE16N 으로 직접 넣었다(ERNAM=ZTEST0911). "
                "테스트 후 모두 원복했다.",
    },
}

io.open(OUT, "w", encoding="utf-8").write(json.dumps(payload, ensure_ascii=False, indent=1))
print("written:", OUT)
res = [r["result"] for r in payload["prog_rows"]]
print("prog rows:", len(res), " PASS:", res.count("PASS"), " FAIL:", res.count("FAIL"), " 미실행:", res.count(""))
cr = [v["result"] for v in payload["common"].values()]
print("common:", len(cr), " Y:", sum(1 for v in payload["common"].values() if v["target"] == "Y"),
      " PASS:", cr.count("PASS"), " FAIL:", cr.count("FAIL"))
print("fail rows:", len(fail_rows), " evidence images:", len(ev_out))
