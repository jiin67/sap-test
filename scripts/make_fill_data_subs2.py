# -*- coding: utf-8 -*-
# make_fill_data_subs2.py : JSON payload for the company unit-test template (v0.02, RFC 수정 후 재테스트)
# 2026-09-13 AS 무상지원금 일시불구독 개선 (ZCRMR058_1 / ZCRMR060_1 / RFC ZCRMF03_RFC_ZCRM_CHRG_CNFRM)
import io, json, os
from PIL import Image as PILImage

OUT = r"C:\sap-test\results\_fill_data_subs2.json"
SHOTS = r"C:\sap-test\results\shots"
THUMB = r"C:\sap-test\results\shots\_tpl_subs2"
os.makedirs(THUMB, exist_ok=True)

DATE = "2026-09-13"
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


prog_header = {
    "B1": "서비스 - AS 무상지원금",
    "D1": "※개발자명 기입 필요",
    "B2": "AS 무상지원금 관리 - 일시불구독 개선 (ZCRM058S / 059S / 060S, 확정일 RFC) [RFC 수정 후 재테스트]",
    "D2": "ZCRMR058_1 / ZCRMR060_1 / ZCRMF03_RFC_ZCRM_CHRG_CNFRM",
}

DATA = ("기간월 2026.07 / X=117786(센터 CT11) 구독 3건 2026070100278·0450·0862, "
        "Y=100011(센터 CT02) 구독 1건 2026070102721")

prog_rows = [
    ("UT-SUB-001", "[1-1] 구독 처리화면 진입 (ZCRM058S)",
     "SUBS_TYPE='S' 인 상담건만 집계되어 표시된다", "Critical",
     DATA + ". 세팅 전 0행 → 세팅 후 X 1행 20,000원(구독 3건). 같은 대리점 일반건 190,000원은 미포함", "PASS"),
    ("UT-SUB-002", "[1-2] 일반 화면 영향 없음 (ZCRM058A)",
     "비구독 건만 표시되고 구독 건은 보이지 않는다", "Critical",
     "X 190,000 → 170,000 (−20,000), Y 65,000 → 60,000 (−5,000). 058A + 058S = 변경 전 058A", "PASS"),
    ("UT-SUB-003", "[1-3] 구독 조회전용 (ZCRM059S)",
     "구독 건이 표시되고 처리 버튼은 표시되지 않는다", "Major",
     "X 1행 20,000, 전건확정 시 확정일 2026.08.07 표시. 툴바는 지급대상데이터조회·상세리스트조회·회계전표출력만. "
     "무상지원금확정/취소, 대리점 확정/확정취소, 일괄회계전표생성/삭제 없음", "PASS"),
    ("UT-SUB-004", "[1-4] 구독 상세 직접 조회 (ZCRM060S)",
     "구독 상담건의 상세만 표시된다", "Major",
     "X: 기사 3명(0000130784 김덕경 5,000 / 0000132094 김태현 10,000 / 0000134384 고대의 5,000), 합계 20,000. "
     "060A 는 기사 8명 합계 170,000", "PASS"),
    ("UT-SUB-005", "[1-5] 상세 드릴다운 (ZCRM058S → 상세리스트조회)",
     "구독 상세(ZCRMR060_1, 구독모드)로 이동한다", "Major",
     "058S X 행 선택 후 상세리스트조회 → ZCRMR060_1 7행, 합계 20,000 (= ZCRM060S 결과와 동일)", "PASS"),

    ("UT-SUB-006", "[2-1] RFC 정상 확정등록",
     "IV_ACTION=S, 유효 상담번호 3건 → EV_RESULT='S', 3건 저장", "Critical",
     "입력 2026070100862:20260807 / 2026070102721:20260804 / 2026070100278:20260805(재전송). "
     "EV_RESULT=S, 3건 모두 '처리완료', ZCRM_CHRG_CNFRM 4행. ※1차(2026-09-11) FAIL → 수정 후 PASS", "PASS"),
    ("UT-SUB-007", "[2-2] RFC 확정취소",
     "IV_ACTION=C → EV_RESULT='S', 해당 건 삭제", "Critical",
     "C / 2026070100862 → EV_RESULT=S, '처리완료', 테이블에서 1건 삭제(3행). "
     "원복 시 4건 일괄 취소도 EV_RESULT=S. ※1차 FAIL → 수정 후 PASS", "PASS"),
    ("UT-SUB-008", "[2-3] RFC 처리구분 누락",
     "IV_ACTION 공백 → 'E', 처리구분 미입력", "Major",
     "※미실행 - SE37 테스트 프레임이 필수 매개변수 IV_ACTION 공백 실행을 막음('매개변수값 입력' 팝업). "
     "소스에는 해당 검증이 있으며, 레거시(CIC) 호출 또는 별도 테스트 프로그램으로 확인 필요", ""),
    ("UT-SUB-009", "[2-4] RFC 잘못된 처리구분",
     "IV_ACTION='X' → 'E', 유효하지 않은 구분", "Major",
     "EV_RESULT=E, ET_RETURN '유효하지 않은 처리구분입니다 (S 또는 C)'. DB 미반영", "PASS"),
    ("UT-SUB-010", "[2-5] RFC 없는 상담번호 포함 (부분성공)",
     "없는 상담번호만 'E', 유효건은 저장되고 EV_RESULT='E'", "Critical",
     "2026070100278:20260805 + 9999999999999:20260805 → 유효건 'S 처리완료'(저장), 없는건 'E 존재하지 않는 상담번호', "
     "전체 EV_RESULT=E. 저장 1건 확인. ※1차 FAIL → 수정 후 PASS", "PASS"),
    ("UT-SUB-011", "[2-6] RFC 확정일 누락",
     "IV_ACTION=S 인데 CFDATE1 공백 → 해당건 'E'(확정일자 누락)", "Major",
     "2026070100450 / 공백 → ET_RETURN CSLTNO 2026070100450(정상 반환), 'E 확정일자 누락'. DB 미반영. "
     "※1차에는 CSLTNO 가 밀려 '존재하지 않는 상담번호' 였음 → 수정 후 PASS", "PASS"),
    ("UT-SUB-012", "[2-7] RFC 미래일자",
     "CFDATE1 미래일자 → 해당건 'E'(미래일자 불가)", "Major",
     "2026070100450 / 2026.12.31 → 'E 미래일자 입력 불가'. DB 미반영. ※1차 FAIL → 수정 후 PASS", "PASS"),
    ("UT-SUB-013", "[2-8] RFC 빈 데이터",
     "IT_DATA 없음 → 'E', 전송 데이터 없음", "Minor",
     "EV_RESULT=E, ET_RETURN '전송 데이터가 없습니다'", "PASS"),
    ("UT-SUB-014", "[2-9] RFC 중복 상담번호",
     "같은 CSLTNO 2줄 → 중복 제거되어 1건 저장", "Major",
     "2026070100450:20260806 두 줄 → EV_RESULT=S, ET_RETURN 은 2줄 모두 '처리완료', "
     "ZCRM_CHRG_CNFRM 에는 1건만 저장. ※1차 FAIL → 수정 후 PASS", "PASS"),

    ("UT-SUB-015", "[3-1] 전건확정 - 미확정 (0/3)",
     "ZCRM058S 대리점확정일(CFDATE1) 공백", "Critical",
     "ZCRM_CHRG_CNFRM 0건 상태에서 058S X CFDATE1 공백", "PASS"),
    ("UT-SUB-016", "[3-2] 전건확정 - 부분확정 (2/3)",
     "3건 중 2건만 확정 → CFDATE1 여전히 공백", "Critical",
     "RFC 로 X-1(2026.08.05) → 1/3, X-2(2026.08.06) → 2/3 확정. 058S X CFDATE1 공백 유지. "
     "※1차에는 RFC 결함으로 SE16N 직접 입력이 필요했으나 이번엔 RFC 로 검증", "PASS"),
    ("UT-SUB-017", "[3-3] 전건확정 - 전건확정 (3/3)",
     "3건 모두 확정 → CFDATE1 = 확정일 최대값", "Critical",
     "+ X-3(2026.08.07) → 058S X CFDATE1 2026.08.07 (08.05/08.06/08.07 중 최대값). 059S 도 동일 표시", "PASS"),
    ("UT-SUB-018", "[3-4] 전건확정 - 확정 후 취소",
     "전건확정에서 1건 취소 → CFDATE1 다시 공백", "Critical",
     "RFC C 로 X-3 취소(3/3 → 2/3) → 058S X CFDATE1 공백", "PASS"),
    ("UT-SUB-019", "[3-5] 전건확정 - 단일건 대리점",
     "대상 1건인 대리점은 1건 확정 즉시 CFDATE1 표시", "Major",
     "Y-1 2026.08.04 확정(SPMON 2026.07) → 058S Y CFDATE1 2026.08.04", "PASS"),

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
     "[예] 선택 시 게이트가 동작하지 않으면 실전표가 생기므로 취소(사용자 결정)", ""),
    ("UT-SUB-024", "[5-1] RFC RSLTDATE 공백 상담건",
     "기간월 산출 불가 오류", "Major",
     "※미실행 - QA 에 조치일자(RSLTDATE)가 공백인 상담건이 없음 (SE16N 조치일자 = 00000000 조회 0건). "
     "소스에는 '결과일자 미존재로 기간월 산출 불가' 검증이 있음", ""),
    ("UT-SUB-025", "[5-2] SPMON 경계 (RSLTDATE 기준 파생)",
     "SPMON 이 상담번호가 아니라 조치일자 기준으로 산출된다", "Major",
     "상담번호 2026062403120(번호상 월 202606 / 조치일자 2026.07.01)을 등록 → SPMON 2026.07 로 저장. "
     "CSLTNO(6) 가정이면 2026.06 이 됐을 값. 확인 후 C 로 즉시 삭제", "PASS"),
    ("UT-SUB-026", "[5-3] 테스트 후 원복",
     "SUBS_TYPE 원복 후 058A 정상 복귀, 구독 화면에 미표시", "Critical",
     "RFC C 로 확정 4건 일괄 취소 → ZCRM_CHRG_CNFRM 0건. SUBS_TYPE 4건 공백 복원. "
     "058A X 190,000 / Y 65,000 (변경 전과 동일), 058S X·Y 0행, 060S X 0행", "PASS"),
    ("UT-SUB-027", "[회귀] 일반 화면에 구독 확정일 미반영",
     "구독 전건확정 상태에서도 ZCRM058A 의 CFDATE1 은 바뀌지 않는다", "Major",
     "3-3 상태(X 3/3 확정)에서 058A X CFDATE1 공백(170,000), 3-5 상태에서 058A Y 공백(60,000)", "PASS"),
    ("UT-SUB-028", "[수정확인] 취소 대상이 없는 건 취소",
     "확정 데이터가 없는 상담건을 C 로 보내면 건별 오류로 처리된다 (전체 롤백 아님)", "Major",
     "이미 취소한 2026070100862 을 C 로 재전송 → EV_RESULT=E, '취소할 확정 데이터가 없습니다'. "
     "EV_RESULT='A'(전체 롤백) 아님", "PASS"),
    ("UT-SUB-029", "[수정확인] 재등록 시 생성이력 유지",
     "이미 있는 확정 건을 다시 등록하면 생성이력은 유지되고 변경이력만 갱신된다", "Minor",
     "2026070100278 을 2-5 에서 등록(생성 23:50:25) 후 2-1 에서 재전송 → 생성 23:50:25 유지, "
     "변경 23:52:44 로 갱신", "PASS"),
]

# ---------------------------------------------------------------- 공통테스트케이스
P = "PASS"
common = {}


def setc(nums, target, result, note=""):
    for n in nums:
        common["UC-%03d" % n] = (target, result, note)


NA = "해당 없음"
# 입력
setc([1], "Y", P, "058 기간월·센터 미입력 시 '모든 필수 입력 필드에 값을 입력하십시오.' 로 실행 차단")
setc([2], "N", "", NA)
setc([3], "Y", P, "RFC 검증 5종(처리구분 공백/오류, 빈 데이터, 상담번호 유효성, 확정일 누락, 미래일자) 정상 동작")
setc([4], "N", "", NA)
setc([5], "Y", P, "미래일자(2026.12.31) 확정일 차단 확인")
setc([6, 7], "N", "", NA)
setc([8], "Y", P, "라디오 기본값(058 '전체', 060 '조회') 확인")
setc([9, 10, 11], "Y", P, "전 실행 구간(058/059/060, SE37 RFC)에서 덤프·행업 없음")
setc([12], "N", "", NA)
setc([13], "Y", P, "검증 실패건은 DB 미반영, 통과분만 저장 (부분성공)")
setc([14], "Y", P, "저장된 값(SPMON/CFDATE1)이 입력값과 일치 - SE16N 확인")
setc([15], "N", "", NA)
setc([16], "Y", P, "동일 CSLTNO 2줄 전송 시 1건만 저장 (중복 제거)")
setc(range(17, 24), "N", "", NA)
# 수정
setc(range(24, 30), "N", "", NA)
setc([30], "Y", P, "기존 확정건 재전송 시 CFDATE1 갱신, 생성이력 유지·변경이력 갱신")
setc(range(31, 40), "N", "", NA)
# 조회
setc([40], "Y", P, "필수 조회조건(기간월·센터) 체크 확인")
setc([41, 42], "N", "", NA)
setc([43], "Y", P, "058S 구독 3건 20,000 = 060S 기사 3명 합계 20,000")
setc([44], "Y", P, "058A + 058S = 변경 전 058A (190,000 = 170,000 + 20,000 / 65,000 = 60,000 + 5,000)")
setc([45], "Y", P, "확정 등록·취소가 058S CFDATE1 에 즉시 반영")
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
setc([61], "Y", P, "취소 대상이 없으면 '취소할 확정 데이터가 없습니다' 로 안내")
setc(range(62, 66), "N", "", NA)
setc([66], "Y", P, "RFC 취소(C) 요청분이 정확히 삭제됨 (1건 취소 / 4건 일괄 취소 모두 확인)")
setc([67], "Y", P, "취소 대상 외 행은 삭제되지 않음 (2/3 상태 유지 확인)")
setc([68, 69], "N", "", NA)
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
setc([99], "Y", P, "RFC 반환 메시지가 사유별로 정확 (확정일자 누락 / 미래일자 입력 불가 / 존재하지 않는 상담번호 구분)")
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
setc([114], "Y", P, "확정일 수신 RFC 정상 동작 - 레거시(CIC) 연계 가능 상태 (1차 결함 해소)")
setc(range(115, 120), "N", "", NA)

assert len(common) == 119, len(common)

fail_rows = [
    {"id": "-", "case": "FAIL 없음", "tester": "-",
     "detail": ("이번 회차(2026-09-12~13, RFC 수정 후) FAIL 은 없다.\n"
                "1차 테스트(2026-09-11)의 FAIL 6건(UT-SUB-006 / 007 / 010 / 011 / 012 / 014)은 "
                "RFC 가 IT_DATA 를 ZCRM_CHRG_CNFRM 형 작업영역으로 옮기면서 CSLTNO 가 3자리 밀리던 결함 때문이었다. "
                "선언부 수정(ls_data / lt_key 를 IT_DATA 라인타입으로) 후 재테스트에서 전부 PASS 로 확인됐다.\n"
                "미실행 3건(UT-SUB-008 2-3 / UT-SUB-022 4-3 / UT-SUB-023 4-4 / UT-SUB-024 5-1)의 사유는 "
                "프로그램별테스트케이스 시트의 테스트 데이터 열에 적었다."),
     "images": []},
]

evidence = [
    ("UT-SUB-001", "구독 세팅 - SE16N 편집 조회 (SUBS_TYPE 공백)", "S13_20260913_0510_2026070100278_before.png"),
    ("UT-SUB-001", "구독 세팅 - 저장 후 (SUBS_TYPE='S')", "S13_20260913_0510_2026070100278_after.png"),
    ("UT-SUB-001", "ZCRM058S 117786 - 구독 3건 20,000원 (전건확정 시 확정일 2026.08.07)", "SUBS_v3_058S_X.png"),
    ("UT-SUB-002", "ZCRM058A 117786 - 170,000원 (구독분 제외)", "SUBS_v3_058A_X.png"),
    ("UT-SUB-002", "ZCRM058A 100011 - 60,000원", "SUBS_v3_058A_Y.png"),
    ("UT-SUB-003", "ZCRM059S 117786 - 처리버튼 없음", "SUBS_v3_059S_X.png"),
    ("UT-SUB-004", "ZCRM060S 117786 - 구독 기사 3명, 합계 20,000", "SUBS_v5_060S_X.png"),
    ("UT-SUB-004", "ZCRM060A 117786 - 일반 기사 8명, 합계 170,000", "SUBS_v4_060A_X.png"),
    ("UT-SUB-005", "058S → 상세리스트조회 드릴다운 결과", "SUBS_v4_drill_X_drill.png"),
    ("UT-SUB-006", "[2-1] RFC 결과화면 - EV_RESULT=S", "S14_R2-1_3_result.png"),
    ("UT-SUB-006", "[2-1] ET_RETURN - 3건 처리완료 (CSLTNO 정상 반환)", "S14_R2-1_4_et_return.png"),
    ("UT-SUB-007", "[2-2] 확정취소 IT_DATA", "S14_R2-2_1_it_data.png"),
    ("UT-SUB-007", "[2-2] 확정취소 ET_RETURN - 처리완료", "S14_R2-2_4_et_return.png"),
    ("UT-SUB-009", "[2-4] 잘못된 처리구분 - 유효하지 않은 처리구분입니다", "S14_R2-4_4_et_return.png"),
    ("UT-SUB-010", "[2-5] 부분성공 - 유효건 처리완료 + 없는건 오류", "S14_R2-5_4_et_return.png"),
    ("UT-SUB-011", "[2-6] 확정일 누락 - CSLTNO 정상 반환", "S14_R2-6_4_et_return.png"),
    ("UT-SUB-012", "[2-7] 미래일자 입력 불가", "S14_R2-7_4_et_return.png"),
    ("UT-SUB-013", "[2-8] 전송 데이터가 없습니다", "S14_R2-8_4_et_return.png"),
    ("UT-SUB-014", "[2-9] 중복 상담번호 - EV_RESULT=S, 1건만 저장", "S14_R2-9_4_et_return.png"),
    ("UT-SUB-015", "[3-1] 확정 0건 - 058S X CFDATE1 공백", "SUBS_v2_31_058S_X.png"),
    ("UT-SUB-016", "[3-2] 부분확정 2/3 - CFDATE1 공백 유지", "SUBS_v2_32_058S_X.png"),
    ("UT-SUB-017", "[3-3] 전건확정 3/3 - CFDATE1 2026.08.07", "SUBS_v2_33_058S_X.png"),
    ("UT-SUB-018", "[3-4] 1건 취소 후 - CFDATE1 다시 공백", "SUBS_v2_34_058S_X.png"),
    ("UT-SUB-019", "[3-5] 단일건 대리점 - CFDATE1 2026.08.04", "SUBS_v2_35_058S_Y.png"),
    ("UT-SUB-023", "[4-4] 일괄회계전표생성 확인 팝업 (취소함, 1차 캡처)", "SUBS_t44_058S_X_fipos_popup1.png"),
    ("UT-SUB-025", "[5-2] SPMON 파생 - 상담번호 202606 이지만 조치일자 기준 2026.07 저장", "S14_R5-2_spmon_4_et_return.png"),
    ("UT-SUB-026", "[5-3] 원복 - 확정 4건 일괄 취소 ET_RETURN", "S14_R5-3_cancel_4_et_return.png"),
    ("UT-SUB-026", "[5-3] 원복 후 058A 117786 - 190,000 (변경 전과 동일)", "SUBS_v5_rb_058A_X.png"),
    ("UT-SUB-026", "[5-3] 원복 후 058S 117786 - 0행", "SUBS_v5_rb_058S_X.png"),
    ("UT-SUB-026", "[5-3] 원복 후 060S 117786 - 0행", "SUBS_v5_rb_060S_X.png"),
    ("UT-SUB-027", "[회귀] 전건확정 상태의 058A 117786 - CFDATE1 공백", "SUBS_v2_reg_058A_X.png"),
    ("UT-SUB-028", "[수정확인] 취소 대상 없음 - 취소할 확정 데이터가 없습니다", "S14_R2-2b_nodata_4_et_return.png"),
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
    "cover": {"C17": "2026.09.13"},
    "revision": {"B5": "0.02", "C5": "RFC 수정 후 재테스트 (1차 FAIL 6건 해소)", "D5": "2026.09.13",
                 "E5": "※작성자 기입 필요", "F5": "", "G5": ""},
    "prog_header": prog_header,
    "prog_rows": [{"id": a, "case": b, "expect": c, "sev": d, "data": e, "result": f}
                  for a, b, c, d, e, f in prog_rows],
    "common": {k: {"target": v[0], "result": v[1], "note": v[2]} for k, v in common.items()},
    "fail_rows": fail_rows,
    "evidence": ev_out,
    "evidence_sheet": "테스트증적(캡처)",
    "legend": {
        "blank": "결과란이 비어 있으면 미실행이다. 사유는 테스트 데이터 열에 기재했다.",
        "note": "대상: ZCRMR058_1 / ZCRMR060_1 / ZCRMF03_RFC_ZCRM_CHRG_CNFRM (KDQ 210). "
                "1차 테스트 2026-09-11 에서 발견한 RFC 결함(CSLTNO 3자리 밀림) 수정 후 2026-09-12~13 재테스트한 회차다. "
                "구독 테스트 데이터는 ZCRM_CICSVC_NEW-SUBS_TYPE 을 4건만 'S' 로 바꿔 만들었고(SE16N), "
                "확정 데이터는 RFC 로 등록·취소했다. 테스트 후 모두 원복했다.",
    },
}

io.open(OUT, "w", encoding="utf-8").write(json.dumps(payload, ensure_ascii=False, indent=1))
print("written:", OUT)
res = [r["result"] for r in payload["prog_rows"]]
print("prog rows:", len(res), " PASS:", res.count("PASS"), " FAIL:", res.count("FAIL"), " 미실행:", res.count(""))
cr = [v["result"] for v in payload["common"].values()]
print("common:", len(cr), " Y:", sum(1 for v in payload["common"].values() if v["target"] == "Y"),
      " PASS:", cr.count("PASS"), " FAIL:", cr.count("FAIL"))
print("evidence images:", len(ev_out), " missing:", missing)
