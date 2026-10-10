# -*- coding: utf-8 -*-
# make_fill_data_194.py : JSON payload for the unit-test result workbook of
#   2026-09-14 scenario "ZCRMR194(무상 매출처리 자동화) -> ZCRM052/052S/053/053S" (SR03000035400)
# consumed by fill_template_subs.ps1 (same template/sheet layout as the 일시불구독 results)
import io, json, os
from PIL import Image as PILImage

OUT = r"C:\sap-test\results\_fill_data_194.json"
SHOTS = r"C:\sap-test\results\shots"
THUMB = r"C:\sap-test\results\shots\_tpl_194"
os.makedirs(THUMB, exist_ok=True)
DATE = "2026-09-14"
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
TD = ("테스트 상담번호 TC1=2026091490001(일시불·센터 CT05) / TC2=2026091490002(구독·대리점 100090, SUBSWRTMONTH 0) / "
      "TC3=2026091490003(구독·대리점 100090, SUBSWRTMONTH 5). 공통: CLMART A, PRICETYPE N, ISFREE Y, INDATE 2026.09.14, SUB COST 0")

# (id, 단계, 테스트 케이스, 예상결과, 심각도, 결과, 테스트 데이터/실제 결과)
CASES = [
    ("P-01", "사전준비", "테스트 데이터 생성 (CICSVC·CICCOST_M·SUB·ADD)",
     "3건 각각 ZCRM_CICSVC / ZCRM_CICCOST_M / ZCRM_CICCOST_SUB / ZCRM_CICCOST_ADD 1행씩 생성, 조건값(CLMART A·ISFREE Y·COST 0·TAGNO·SUBS_TYPE) 반영",
     "Major", P,
     TD + ". SE16N 으로 12행 삽입(각 '삽입된 행 1'), 생성자 ZTEST0914. 기존 정상건 2026062600687(TC1)·2026063003075(TC2·TC3) 값을 복사하고 "
     "고객명·주소·전화는 ZTEST 값으로 대체. ※SUBSWRTMONTH 에 음수(-3, 3-)를 넣으면 SE16N 이 오류 없이 저장하지 않아 시나리오 허용값 0 으로 생성"),
    ("A-01", "STEP4", "ZCRMR194 실행 (대상 추출·실행 결과)",
     "S_VKORG ASKR, S_CSLTNO 테스트 3건, S_INDAT 오늘 → 리스트 '청구 저장/확정완료', '전체 : 3'",
     "Critical", P,
     "SE38 ZCRMR194 실행(T-code 없음) → 리스트 '청구 저장/확정완료' / '전체 : 3' / '에러 혹은 유상 : 2'. 조건 밖 데이터는 처리되지 않음"),
    ("A-02", "STEP4", "ZCRM_REQ 요청번호 생성",
     "3건 REQNO 생성, REQNO_ITEM 10/20/30…", "Critical", P,
     "REQNO 2026091400001 / REQNO_ITEM 10(TC1)·20(TC2)·30(TC3), CSLTNO_ITEM 1"),
    ("A-03", "STEP4", "청구확정 REQ_CONF",
     "3건 REQ_CONF = 'X'", "Critical", P, "3건 모두 REQ_CONF X, REQ_DROP·GIDROP 공백"),
    ("A-04", "STEP4", "청구요청일 PRDREQDATE",
     "PRDREQDATE = 실행일(sy-datum)", "Major", P, "3건 모두 2026.09.14"),
    ("A-05", "STEP4", "꼬리표 TAGNO1 복사",
     "TAGNO1 = ZCRM_CICSVC-TAGNO, TAGNO1_CHECK = 'X'", "Critical", P,
     "TAGNO 9141/9142/9143 → TAGNO1 9141/9142/9143, TAGNO1_CHECK X. REQ_PERSON = SVCNAME(대전센터 / (주)평택경동(평택대리점))"),
    ("A-06", "STEP4", "요청번호 채번 ZCRM_REQNO",
     "오늘자 채번 순번 증가", "Minor", P, "ZCRM_REQNO CURRDT 2026.09.14 / REQNO 1 → 요청번호 20260914 + 00001"),
    ("A-07", "STEP4", "무상매출 자동처리 (ISFREE='Y')",
     "가용재고 있는 건: 무상오더 ZFD1 → 납품 → 출고, STATUS1~3_FD 녹색, VBELN_VA_FD/VBELN_VL_FD 기록. 실패 시 ZCRM_AUTO_ERRLOG",
     "Major", P,
     "TC2: 오더 4293318(ZFD1, PO번호 2026091490002) → 납품 85857197(ZLFD, 실제출고일 2026.09.14) → 자재전표 4921917570, STATUS1~3_FD 녹색, ZCRM_ORDER_AUTO 3건. "
     "TC1·TC3: ZCRM_AUTO_ERRLOG MAKE_ORDER '자재 30037654A / 30031356A 가용수량이 부족합니다' 로 매출 제외(가용재고 부족 - TC2 가 CT04 재고 1개를 먼저 출고). ZCRM_REQ 는 3건 모두 생성"),
    ("A-08", "STEP4", "에러 건수 표시",
     "'에러 혹은 유상' 건수 = ZCRM_AUTO_ERRLOG 생성 건수", "Minor", P, "리스트 2 = ERRLOG 2건(TC1, TC3)"),
    ("A-09", "STEP4", "(역검증) ISFREE≠'Y' 유무상 혼재",
     "type 'C' → '유무상 혼재 건' 에러로그, 매출 제외", "Minor", N,
     "※미실행 - 사용자 결정으로 3건 모두 ISFREE='Y' 로 생성(시나리오 선택 항목)"),
    ("B-01", "STEP5", "ZCRM052 TC1 조회",
     "센터조치 + 회수품 미접수 + 상담번호 TC1 → 1건", "Critical", P,
     "상담번호 2026091490001 → '회수품적합성관리 (1)'. 범위 TC1~TC3 로 조회해도 TC1 1건"),
    ("B-02", "STEP5", "유무상 표기 (비구독)",
     "PRICETYPET = '무상'", "Critical", P, "서비스 처리결과 '무상'"),
    ("B-03", "STEP5", "구독구분 컬럼 미표시",
     "비구독 T-code 는 SUBS_TYPE 미표시(NO_OUT)", "Minor", P, "표시 컬럼 42개에 구독구분 없음, 세부사항 팝업에도 없음"),
    ("B-04", "STEP5", "화면 제목 (T000)",
     "기존 제목 유지", "Minor", P, "선택화면 '회수품적합성관리', 결과화면 '회수품적합성관리 (1)', GUI 상태 S000"),
    ("B-05", "STEP5", "구독건 제외 (subs_type <> 'S')",
     "ZCRM052 에서 TC2·TC3 미조회", "Critical", P,
     "서비스점조치 + 미접수 + 상담번호 TC1~TC3 → '(0)' (TC2·TC3 은 RESULTORG B 이지만 구독이라 제외)"),
    ("B-06", "STEP5", "미접수 조회 시 접수일자 조건",
     "회수품 미접수 선택 시 접수일자(SO_RCDT) 미사용", "Minor", P, "미접수 선택 시 SO_RCDT 필드 숨김, 청구요청일 공백으로 조회됨"),
    ("C-01", "STEP6", "ZCRM052S 선택화면",
     "센터조치(PA_R01) 숨김, 대리점조치(PA_R02)만 표시·선택, 블록 제목 '구독조치', 구독 전용 제목(T52S)", "Minor", P,
     "radPA_R01 없음, radPA_R02 선택됨, 블록 '구독조치', 제목 '[구독전용] 회수품 적합성 관리'"),
    ("C-02", "STEP6", "ZCRM052S 구독건 조회",
     "미접수 + 상담번호 TC1~TC3 → TC2·TC3 2건", "Critical", P, "2건(2026091490002, 2026091490003)"),
    ("C-03", "STEP6", "TC2 유무상 표기",
     "SUBSWRTMONTH ≤ 0 → '구독무상(N년이내)'", "Critical", P, "SUBSWRTMONTH 0(경계값) → '구독무상(N년이내)'"),
    ("C-04", "STEP6", "TC3 유무상 표기",
     "SUBSWRTMONTH > 0 → '구독무상(N년이후)'", "Critical", P, "SUBSWRTMONTH 5 → '구독무상(N년이후)'"),
    ("C-05", "STEP6", "구독구분 컬럼 표시",
     "ZCRM052S 결과 ALV 에 구독구분(SUBS_TYPE) 컬럼 표시", "Minor", F,
     "필드카탈로그에는 있으나(NO_OUT 해제) 기본 레이아웃에서 '숨긴 열'로 빠져 화면에 안 보임 → FAIL내역 참조"),
    ("C-06", "STEP6", "일시불건 제외",
     "ZCRM052S 에서 TC1 미조회", "Critical", P, "TC1 조회 안 됨"),
    ("C-07", "STEP6", "(추가 발견) 결과화면 제목 조회건수",
     "기존처럼 제목에 조회건수 '(n)' 표시", "Minor", F,
     "ZCRM052/053 은 '(1)' 표시, ZCRM052S 는 '[구독전용] 회수품 적합성 관리' 만 표시(건수 없음) → FAIL내역 참조"),
    ("D-01", "STEP7", "ZCRM053 조회전용",
     "TC1 조회, 편집불가, [저장] 제외, 제목 T100", "Major", P,
     "1건 '무상', 제목 '회수품적합성관리 조회 (1)', 저장 버튼 비활성. 편집 가능 셀은 회수품입고일자 1개(기존과 동일, 저장 불가)"),
    ("D-02", "STEP7", "ZCRM053S 조회전용(구독)",
     "TC2·TC3 조회, 편집불가, [저장] 제외, 구독표기, 제목 T53S", "Critical", F,
     "T-code 실행 즉시 ABAP 덤프 LOAD_PROGRAM_NOT_FOUND (Program 'ZCRMR053' not found). TSTC 에 ZCRM053S → ZCRMR053 로 등록됨 → FAIL내역 참조"),
    ("E-01", "STEP8", "구독건 회수품 접수 저장",
     "적합수량·입고일자 입력 후 [저장]→'예' → 정상 저장, ZCRM_RECPT 접수번호 생성(RECPTDATE=오늘)", "Critical", P,
     "ZCRM052S 미접수 TC2: 적합수량 1 / 책임사업장 1100 / 회수품입고일자 2026.09.14 / 접수부서·접수자 ZTEST_20260914 → 확인 팝업 '예' → "
     "'정상적으로 저장되었습니다.' ZCRM_RECPT 접수번호 2026091400001-10, 접수일자 2026.09.14, NCF Y, STATUS X, 금액 39,500 KRW"),
    ("E-02", "STEP8", "접수 조회 (PA_CHX)",
     "회수품 접수 + 접수일자 오늘 → 저장한 TC2 조회", "Major", P,
     "접수일자 2026.09.14 → TC2 1건, 접수번호 2026091400001, 적합구분 '적합', '구독무상(N년이내)'. 미접수 목록에서는 TC2 빠짐(TC3 만 남음)"),
    ("E-03", "STEP8", "구독 검증 CHECK_SUBSCRIPTION_DATA",
     "구독 아닌 건이 섞이면 '구독 대상이 아닌 데이터가 포함되어 있습니다.' 후 저장 중단", "Major", N,
     "※미실행 - ZCRM052S 조회 자체가 SUBS_TYPE='S' 만 가져오므로 화면에서 혼재 상황을 만들 수 없음. 소스로 저장 직전 검증 호출만 확인"),
    ("G-01", "회귀", "ZCRM052 기존 데이터 조회",
     "구독 분기 추가 후에도 일반 조회 건수 유지", "Critical", P,
     "청구요청일 2026.07.01~07.15 미접수: 센터조치 391건(09-10 테스트 395건 - 그날 접수 저장한 4건), 서비스점조치 619건, ZCRM052S 0건(해당 기간 구독 데이터 없음)"),
]
prog_rows = [{"id": c[0], "case": "[%s] %s" % (c[1], c[2]), "expect": c[3], "sev": c[4], "result": c[5], "data": c[6]} for c in CASES]

# ---------------- 공통테스트케이스 ----------------
common = {}


def setc(nums, target, result, note=""):
    for n in nums:
        common["UC-%03d" % n] = (target, result, note)


NA = "해당 없음"
setc(range(1, 120), "N", "", NA)
setc([3], "Y", P, "ZCRMR194 대상 검증(CLMART A·교체품·가용재고) → 조건 미충족 건 에러로그 후 매출 제외")
setc([8], "Y", P, "ZCRM052S 서비스점조치 기본선택, 접수일자 기본값 오늘")
setc([9, 10, 11], "Y", P, "ZCRMR194·ZCRM052/052S 실행·저장 중 덤프/행업 없음")
setc([13], "Y", P, "ZCRM_REQ 3건·ZCRM_RECPT 1건 누락 없이 저장")
setc([14], "Y", P, "TAGNO1=TAGNO, PRDREQDATE=오늘, 적합수량·입고일자 입력값대로 저장")
setc([16], "Y", "", "미실행 - ZCRMR194 재실행 시 중복 방지 미검증")
setc([22], "Y", P, "요청번호 20260914+00001, 품목 10/20/30 / 접수번호 2026091400001")
setc([40], "Y", P, "조회조건(처리조직·접수여부·상담번호) 정상")
setc([43], "Y", P, "TC별 조회 건수 기대값과 일치, 회귀 391건 정합")
setc([45], "Y", P, "생성·저장 데이터가 052/052S/053 조회에 반영")
setc([46], "Y", P, "ALV 세부사항 팝업 정상")
setc([49], "Y", F, "ZCRM053S 실행 시 덤프 LOAD_PROGRAM_NOT_FOUND → D-02")
setc([50, 51], "Y", P, "대기·멈춤 없음")
setc([54], "Y", P, "처리조직↔센터/서비스점, 접수여부↔접수일자 필드 연동")
setc([55], "Y", P, "상담번호 + 처리조직 + 접수여부 + 구독구분 조건 동시 적용")
setc([57], "Y", P, "상담번호당 1행")
setc([87], "Y", P, "개발용 메시지 노출 없음")
setc([88], "Y", P, "053 저장 버튼 비활성, 052S 편집 셀 제어")
setc([93], "Y", F, "ZCRM052S 구독구분 컬럼 미표시(C-05), 결과 제목 건수 미표시(C-07)")
setc([94], "Y", F, "T-code ZCRM053S 가 없는 프로그램 ZCRMR053 에 연결(D-02)")
setc([97], "Y", P, "저장 확인 팝업 예/아니오/취소")
setc([99], "Y", P, "'정상적으로 저장되었습니다.', '청구 저장/확정완료' 메시지")
setc([101], "Y", P, "라디오 초기값(052 센터조치 / 052S 서비스점조치)")
setc([106], "Y", P, "T-code 별 SAVE 제외·라디오 숨김")
setc([109], "Y", P, "053 편집불가, 052S 적합수량·책임사업장·입고일자 편집가능")
setc([111], "Y", P, "한글 메시지")
setc([114], "Y", F, "ZCRM053S 실행 불가(D-02)")
assert len(common) == 119, len(common)

fail_rows = [
    {"id": "D-02", "case": "[STEP7] ZCRM053S 조회전용(구독) - 실행 불가", "tester": "※테스트 수행자 기입 필요",
     "detail": ("[현상] /nZCRM053S 실행 즉시 ABAP 실행오류(덤프) LOAD_PROGRAM_NOT_FOUND - Program \"ZCRMR053\" not found.\n"
                "[원인] TSTC 에 ZCRM053S 의 프로그램이 ZCRMR053 으로 등록되어 있다(ZCRM052·ZCRM052S·ZCRM053 은 ZCRMR052). "
                "KDQ 에 ZCRMR053 소스가 없다(소스스캔 결과 없음). ZCRMR052 소스에는 sy-tcode = 'ZCRM053S' 분기(구독·조회전용, 제목 T53S)가 구현돼 있다.\n"
                "[조치 제안] SE93 에서 ZCRM053S 의 프로그램을 ZCRMR052(화면 1000)로 수정 후 재테스트. STEP7-2(편집불가·저장 제외·구독표기·T53S)는 미검증 상태."),
     "images": ["S20_s7b_1_selection.png", "S12_TSTC_e7_tstc.png"]},
    {"id": "C-05", "case": "[STEP6] ZCRM052S 구독구분 컬럼 미표시", "tester": "※테스트 수행자 기입 필요",
     "detail": ("[현상] ZCRM052S 결과 ALV 표시 컬럼 42개에 '구독구분'이 없다. 레이아웃 변경 팝업을 열면 '구독구분'이 숨긴 열 목록에 있다(값은 세부사항 팝업에서 S 로 확인됨).\n"
                "[원인 추정] 필드카탈로그는 구독모드에서 NO_OUT 을 해제하지만, 레이아웃 키(GV_VARIANT-REPORT = SY-REPID+화면+PA_CONS_NAME)가 052/052S/053/053S 공통이라 "
                "기존에 저장된 기본 레이아웃(변경 전 저장, 구독구분 없음)이 적용되어 새 컬럼이 숨겨진다.\n"
                "[조치 제안] 구독 T-code 용 레이아웃 HANDLE 분리 또는 기본 레이아웃에 구독구분 추가."),
     "images": ["S20_s6a_3_result.png", "S20_s6lay_layout.png"]},
    {"id": "C-07", "case": "[STEP6 추가 발견] ZCRM052S 결과화면 제목에 조회건수 미표시", "tester": "※테스트 수행자 기입 필요",
     "detail": ("[현상] ZCRM052 '회수품적합성관리 (1)', ZCRM053 '회수품적합성관리 조회 (1)' 은 건수를 보여주나, "
                "ZCRM052S 는 2건 조회해도 '[구독전용] 회수품 적합성 관리' 로만 표시된다.\n"
                "[원인 추정] STATUS_0100 에서 set titlebar 'T52S' with gv_list_count_text 로 넘기지만 제목 T52S 텍스트에 &1 자리표시자가 없는 것으로 보인다(T53S 도 확인 필요).\n"
                "[조치 제안] GUI 제목 T52S/T53S 텍스트에 &1 추가."),
     "images": ["S20_s6a_3_result.png", "S20_s5a_2_result.png"]},
]

evidence = [
    ("P-01", "테스트 데이터 - ZCRM_CICSVC 입력 화면(TC1, 저장 전)", "S19_tc1svc_ZCRM_CICSVC_2026091490001_before_save.png"),
    ("P-01", "SE16N 저장 결과 '삽입된 행 1'", "S19_tc1svc_ZCRM_CICSVC_2026091490001_saved.png"),
    ("P-01", "ZCRM_CICSVC 3건 - 꼬리표·클레임유형", "S12_ZCRM_CICSVC_e0_svc_tag.png"),
    ("P-01", "ZCRM_CICSVC 3건 - 조치영역(RESULTORG)·서비스점", "S12_ZCRM_CICSVC_e0_svc_org.png"),
    ("P-01", "ZCRM_CICSVC 3건 - 구독여부 S / 구독무상 개월수 0·0·5", "S12_ZCRM_CICSVC_e0_svc_subs.png"),
    ("P-01", "ZCRM_CICCOST_M 3건 - 유무상여부 Y, 거래일자 2026.09.14", "S12_ZCRM_CICCOST_M_e0_costm_free.png"),
    ("P-01", "ZCRM_CICCOST_SUB 3건 - 교체자재, 수량 1, 표준금액 0", "S12_ZCRM_CICCOST_SUB_e0_sub.png"),
    ("A-01", "ZCRMR194 선택화면 (ASKR / 상담번호 TC1~TC3 / 거래일자 오늘)", "S20_r194run_1_selection.png"),
    ("A-01", "실행 결과 - '청구 저장/확정완료', 전체 3, 에러 혹은 유상 2", "S20_r194run_2_result.png"),
    ("A-02", "ZCRM_REQ - 요청번호 2026091400001 품목 10/20/30, 청구요청확인 X, 꼬리표 9141~9143 (A-03~A-05)", "S12_ZCRM_REQ_e194_req.png"),
    ("A-06", "ZCRM_REQNO - 2026.09.14 / 1", "S12_ZCRM_REQNO_e194_reqno.png"),
    ("A-07", "ZCRM_CICCOST_M - TC2 오더/납품/출고(무상) 녹색", "S12_ZCRM_CICCOST_M_e194_costm.png"),
    ("A-07", "ZCRM_CICCOST_M - TC2 판매문서 4293318 / 납품 85857197", "S12_ZCRM_CICCOST_M_e194_costm2.png"),
    ("A-07", "ZCRM_ORDER_AUTO - TC2 오더·납품·자재전표 3건", "S12_ZCRM_ORDER_AUTO_e194_auto.png"),
    ("A-07", "VBAK 4293318 - 오더유형 ZFD1, PO번호 2026091490002", "S12_VBAK_e194_vbak.png"),
    ("A-07", "LIKP 85857197 - 납품유형 ZLFD, 실제출고일 2026.09.14", "S12_LIKP_e194_likp.png"),
    ("A-08", "ZCRM_AUTO_ERRLOG - TC1·TC3 '가용수량이 부족합니다'", "S12_ZCRM_AUTO_ERRLOG_e194_err.png"),
    ("B-01", "ZCRM052 선택화면 - 센터조치 + 미접수, 접수일자 숨김 (B-06)", "S20_s5a_1_selection.png"),
    ("B-01", "ZCRM052 결과 - TC1 1건, 제목 '회수품적합성관리 (1)' (B-04)", "S20_s5a_2_result.png"),
    ("B-02", "ZCRM052 TC1 세부사항 - 서비스 처리결과 '무상', 구독구분 항목 없음 (B-03)", "S20_d5_detail_row0.png"),
    ("B-04", "ZCRM052 시스템 > 상태 - GUI 상태 S000", "S20_s5a_status_12.png"),
    ("B-05", "ZCRM052 서비스점조치 + 상담번호 TC1~TC3 → (0) : 구독건 제외", "S20_s5b_2_result.png"),
    ("C-01", "ZCRM052S 선택화면 - '구독조치' 블록, 서비스점조치만 표시", "S20_s6a_1_selection.png"),
    ("C-02", "ZCRM052S 결과 - TC2·TC3 2건 (TC1 없음, C-06)", "S20_s6a_3_result.png"),
    ("C-04", "ZCRM052S TC3 세부사항 - 구독구분 S, '구독무상(N년이후)'", "S20_d6_detail_row0.png"),
    ("C-03", "ZCRM052S TC2 세부사항 - '구독무상(N년이내)' (STEP8 접수 후 접수조회 화면)", "S20_d8_detail_row0_b.png"),
    ("C-05", "ZCRM052S 레이아웃 변경 - 구독구분이 숨긴 열에 있음 (결함)", "S20_s6lay_layout.png"),
    ("D-01", "ZCRM053 결과 - '회수품적합성관리 조회 (1)', 저장 버튼 비활성", "S20_s7a_2_result.png"),
    ("D-01", "ZCRM053 TC1 세부사항 - '무상'", "S20_d7_detail_row0.png"),
    ("D-02", "ZCRM053S 실행 - 덤프 LOAD_PROGRAM_NOT_FOUND (결함)", "S20_s7b_1_selection.png"),
    ("D-02", "TSTC - ZCRM053S → ZCRMR053 (결함 원인)", "S12_TSTC_e7_tstc.png"),
    ("E-01", "ZCRM052S TC2 입력 - 적합수량 1, 책임사업장 1100, 입고일자, 접수부서/접수자", "S20_s8a_1_entered.png"),
    ("E-01", "저장 확인 팝업", "S20_s8a_popup_w1_18.png"),
    ("E-01", "저장 완료 - '정상적으로 저장되었습니다.'", "S20_s8a_2_after_save.png"),
    ("E-01", "ZCRM_RECPT - 접수번호 2026091400001, 접수일자 2026.09.14, 적합수량 1, NCF Y", "S12_ZCRM_RECPT_e194_recpt.png"),
    ("E-02", "ZCRM052S 접수 조회 - TC2 1건", "S20_s8b_2_result.png"),
    ("E-02", "TC2 세부사항 - 접수번호·접수일자·적합수량", "S20_d8_detail_row0.png"),
    ("G-01", "회귀 - ZCRM052 센터조치 미접수 2026.07.01~15 (391)", "S20_g1a_result.png"),
    ("G-01", "회귀 - ZCRM052 서비스점조치 미접수 (619)", "S20_g1c_result.png"),
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
    "cover": {"C17": "2026.09.14"},
    "revision": {"B5": "0.01", "C5": "ZCRMR194 → ZCRM052/052S/053/053S 무상 청구·구독 회수품 end-to-end 테스트 최초 작성", "D5": "2026.09.14",
                 "E5": "※작성자 기입 필요", "F5": "", "G5": ""},
    "prog_header": {
        "B1": "서비스 - 회수품 관리",
        "D1": "※개발자명 기입 필요",
        "B2": "SR03000035400 보일러 서비스 구독 대리점 비용 정산 - 무상 매출처리 자동화(ZCRMR194) → 회수품 적합성 관리(ZCRM052/052S/053/053S)",
        "D2": "ZCRMR194 / ZCRMR052",
    },
    "prog_rows": prog_rows,
    "common": {k: {"target": v[0], "result": v[1], "note": v[2]} for k, v in common.items()},
    "fail_rows": fail_rows,
    "evidence": ev_out,
    "evidence_sheet": "테스트증적(캡처)",
    "legend": {
        "blank": "결과란이 비어 있으면 미실행이다(사유는 테스트 데이터 열). PASS 건의 성공 근거 화면은 '테스트증적(캡처)' 시트에 케이스 ID 별로 넣었다.",
        "note": "KDQ 210, 2026-09-14. 테스트 데이터(상담번호 2026091490001~003, 생성자 ZTEST0914)와 ZCRMR194·052S 가 만든 데이터(요청번호·오더·납품·출고·접수번호)는 "
                "사용자 결정에 따라 삭제하지 않고 KDQ 에 남겨 두었다. 상세는 '테스트데이터' 시트.",
    },
}
io.open(OUT, "w", encoding="utf-8").write(json.dumps(payload, ensure_ascii=False, indent=1))
res = [r["result"] for r in prog_rows]
print("prog rows:", len(res), " PASS:", res.count("PASS"), " FAIL:", res.count("FAIL"), " 미실행:", res.count(""))
cr = payload["common"].values()
print("common Y:", sum(1 for v in cr if v["target"] == "Y"), " PASS:", sum(1 for v in cr if v["result"] == "PASS"), " FAIL:", sum(1 for v in cr if v["result"] == "FAIL"))
print("evidence:", len(ev_out), " missing:", missing)
