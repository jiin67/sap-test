# -*- coding: utf-8 -*-
# make_fill_data_194b.py : payload for the 2026-09-18 RETEST of "ZCRMR194 → ZCRM052/052S/053/053S" (v0.02).
#   Reuses the common-checklist block of make_fill_data_194.py (v0.01) with the retest results patched in.
import io, json, os, re
from PIL import Image as PILImage

OUT = r"C:\sap-test\results\_fill_data_194b.json"
SHOTS = r"C:\sap-test\results\shots"
THUMB = r"C:\sap-test\results\shots\_tpl_194b"
os.makedirs(THUMB, exist_ok=True)
DATE = "2026-09-18"
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
TD = ("테스트 상담번호 TC1=2026091891001(일시불·센터 CT05) / TC2=2026091891002(구독·대리점 100090, SUBSWRTMONTH 0) / "
      "TC3=2026091891003(구독·대리점 100090, SUBSWRTMONTH 5). 공통: CLMART A, PRICETYPE N, ISFREE Y, INDATE 2026.09.18, SUB COST 0")
FIX = "[재테스트] "
CASES = [
    ("P-01", "사전준비", "테스트 데이터 생성 (CICSVC·CICCOST_M·SUB·ADD)", "3건 각각 4개 테이블 1행씩 생성, 조건값 반영", "Major", P,
     TD + ". SE16N 으로 12행 삽입(각 '삽입된 행 1'), 생성자 ZTEST0918B. 복사 원본은 1차와 동일(2026062600687 / 2026063003075). ※SUBSWRTMONTH 음수는 SE16N 으로 저장되지 않아 TC2 는 0 사용(1차와 동일)"),
    ("A-01", "STEP4", "ZCRMR194 실행", "'청구 저장/확정완료', '전체 : 3'", "Critical", P, "SE38 ZCRMR194 → '청구 저장/확정완료' / '전체 : 3' / '에러 혹은 유상 : 3'"),
    ("A-02", "STEP4", "ZCRM_REQ 요청번호 생성", "3건 REQNO 생성, 품목 10/20/30", "Critical", P, "REQNO 2026091800001 / 품목 10(TC1)·20(TC2)·30(TC3)"),
    ("A-03", "STEP4", "청구확정 REQ_CONF", "3건 'X'", "Critical", P, "3건 모두 X, REQ_DROP 공백"),
    ("A-04", "STEP4", "청구요청일 PRDREQDATE", "실행일", "Major", P, "3건 모두 2026.09.18"),
    ("A-05", "STEP4", "꼬리표 TAGNO1 복사", "TAGNO1 = CICSVC-TAGNO, CHECK 'X'", "Critical", P, "9281/9282/9283 일치, TAGNO1_CHECK X, REQ_PERSON = SVCNAME"),
    ("A-06", "STEP4", "요청번호 채번 ZCRM_REQNO", "오늘자 순번 증가", "Minor", P, "ZCRM_REQNO 2026.09.18 / 1"),
    ("A-07", "STEP4", "무상매출 자동처리 (ISFREE='Y')", "가용재고 있으면 ZFD1 오더→납품→출고, 없으면 ZCRM_AUTO_ERRLOG", "Major", P,
     "이번에는 3건 모두 가용재고 없음 → ZCRM_AUTO_ERRLOG MAKE_ORDER '자재 30037654A / 30031356A 가용수량이 부족합니다', 오더 미생성, 상태 공백. 청구(ZCRM_REQ)는 3건 모두 생성. "
     "※오더·납품·출고 정상 경로는 1차(09-14, 오더 4293318)에서 확인. CT04 의 30031356A 재고 1개는 그때 출고됨"),
    ("A-08", "STEP4", "에러 건수 표시", "리스트 건수 = ERRLOG 건수", "Minor", P, "리스트 3 = ERRLOG 3건"),
    ("A-09", "STEP4", "(역검증) ISFREE≠'Y' 유무상 혼재", "'유무상 혼재 건' 에러로그", "Minor", N, "※미실행 - 1차와 같이 3건 모두 ISFREE='Y' 로 생성(선택 항목)"),
    ("B-01", "STEP5", "ZCRM052 TC1 조회", "센터조치 + 미접수 → 1건", "Critical", P, "상담번호 TC1~TC3 범위 → '회수품적합성관리 (1)' TC1 만 조회"),
    ("B-02", "STEP5", "유무상 표기 (비구독)", "'무상'", "Critical", P, "서비스 처리결과 '무상'"),
    ("B-03", "STEP5", "구독구분 컬럼 미표시", "비구독 T-code 는 구독구분 컬럼 미표시(NO_OUT)", "Minor", F,
     "[신규] ZCRM052·ZCRM053 결과 ALV 맨 앞에 '구독'(구독구분) 컬럼이 표시됨(값은 공백). 1차에서는 표시되지 않았음 → FAIL내역 참조"),
    ("B-04", "STEP5", "화면 제목 (T000)", "기존 제목", "Minor", P, "'회수품적합성관리' / '회수품적합성관리 (1)'"),
    ("B-05", "STEP5", "구독건 제외", "ZCRM052 에서 TC2·TC3 미조회", "Critical", P, "서비스점조치 + 미접수 + TC1~TC3 → '(0)'"),
    ("B-06", "STEP5", "미접수 조회 시 접수일자 조건", "접수일자 미사용", "Minor", P, "미접수 선택 시 접수일자 필드 숨김"),
    ("C-01", "STEP6", "ZCRM052S 선택화면", "센터조치 숨김, 대리점조치만, '구독조치', 구독 전용 제목", "Minor", P, "radPA_R01 없음, radPA_R02 선택, 블록 '구독조치', 제목 '[구독전용] 회수품 적합성 관리'"),
    ("C-02", "STEP6", "ZCRM052S 구독건 조회", "TC2·TC3 2건", "Critical", P, "2건(2026091891002, 2026091891003)"),
    ("C-03", "STEP6", "TC2 유무상 표기", "'구독무상(N년이내)'", "Critical", P, "SUBSWRTMONTH 0 → '구독무상(N년이내)'"),
    ("C-04", "STEP6", "TC3 유무상 표기", "'구독무상(N년이후)'", "Critical", P, "SUBSWRTMONTH 5 → '구독무상(N년이후)'"),
    ("C-05", "STEP6", "구독구분 컬럼 표시", "ZCRM052S 결과 ALV 에 구독구분 표시", "Minor", P, FIX + "1차 FAIL → 수정 확인. 결과 ALV 첫 컬럼에 '구독' = S 표시(서비스 처리결과도 앞쪽으로 이동)"),
    ("C-06", "STEP6", "일시불건 제외", "TC1 미조회", "Critical", P, "TC1 조회 안 됨"),
    ("C-07", "STEP6", "결과화면 제목 조회건수", "제목에 '(n)' 표시", "Minor", P, FIX + "1차 FAIL → 수정 확인. '[구독전용] 회수품 적합성 관리 (2)'"),
    ("D-01", "STEP7", "ZCRM053 조회전용", "TC1 조회, 편집불가, 저장 제외, T100", "Major", P, "1건 '무상', '회수품적합성관리 조회 (1)', 저장 버튼 비활성, 편집 가능 셀은 회수품입고일자 1개(기존 동일, 저장 불가)"),
    ("D-02", "STEP7", "ZCRM053S 조회전용(구독)", "TC2·TC3 조회, 편집불가, 저장 제외, 구독표기, T53S", "Critical", P,
     FIX + "1차 FAIL(덤프) → 수정 확인. TSTC ZCRM053S → ZCRMR052. 선택화면 '[구독전용] 회수품 적합성 관리 조회', 센터조치 숨김·'구독조치'. "
     "결과 '… 조회 (2)', TC2 '구독무상(N년이내)'·TC3 '구독무상(N년이후)', 구독 S, 저장 버튼 비활성, 편집 가능 셀은 회수품입고일자만"),
    ("E-01", "STEP8", "구독건 회수품 접수 저장", "정상 저장, 접수번호 생성", "Critical", P, "ZCRM052S TC2: 적합수량 1 / 1100 / 입고일자 2026.09.18 → '정상적으로 저장되었습니다.' 접수번호 2026091800014-10, NCF Y, 39,500 KRW"),
    ("E-02", "STEP8", "접수 조회 (PA_CHX)", "접수일자 오늘 → TC2 조회", "Major", P, "TC2 1건, 접수번호 2026091800014, '적합', '구독무상(N년이내)', 제목 '(1)'"),
    ("E-03", "STEP8", "구독 검증 CHECK_SUBSCRIPTION_DATA", "비구독 혼재 시 저장 중단", "Major", N, "※미실행 - ZCRM052S 는 구독건만 조회되어 화면에서 혼재 상황을 만들 수 없음(1차와 동일)"),
    ("G-01", "회귀", "ZCRM052 기존 데이터 조회", "일반 조회 건수 유지", "Critical", P, "청구요청일 2026.07.01~07.15 미접수: 센터조치 391건, 서비스점조치 619건 (1차와 동일)"),
]
prog_rows = [{"id": c[0], "case": "[%s] %s" % (c[1], c[2]), "expect": c[3], "sev": c[4], "result": c[5], "data": c[6]} for c in CASES]

# ---- common checklist: take the v0.01 block and patch the retest results
src = io.open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "make_fill_data_194.py"), encoding="utf-8").read()
blk = src[src.index("common = {}"):src.index("fail_rows = [")]
blk = blk.replace('setc([49], "Y", F, "ZCRM053S 실행 시 덤프 LOAD_PROGRAM_NOT_FOUND → D-02")', 'setc([49], "Y", P, "ZCRM053S 덤프 해소, 전 구간 덤프 없음")')
blk = blk.replace('setc([93], "Y", F, "ZCRM052S 구독구분 컬럼 미표시(C-05), 결과 제목 건수 미표시(C-07)")',
                  'setc([93], "Y", F, "ZCRM052·053(비구독) 결과에 구독구분 컬럼이 표시됨(B-03). 1차 지적 C-05·C-07 은 수정 확인")')
blk = blk.replace('setc([94], "Y", F, "T-code ZCRM053S 가 없는 프로그램 ZCRMR053 에 연결(D-02)")', 'setc([94], "Y", P, "ZCRM053S → ZCRMR052 로 수정됨")')
blk = blk.replace('setc([114], "Y", F, "ZCRM053S 실행 불가(D-02)")', 'setc([114], "Y", P, "4개 T-code 모두 실행 가능")')
blk = blk.replace("요청번호 20260914+00001, 품목 10/20/30 / 접수번호 2026091400001", "요청번호 20260918+00001, 품목 10/20/30 / 접수번호 2026091800014")
assert blk.count(', F, ') == 1, blk.count(', F, ')
exec(blk)

fail_rows = [
    {"id": "B-03", "case": "[STEP5 신규] ZCRM052·ZCRM053(비구독) 결과에 구독구분 컬럼이 표시됨", "tester": "※테스트 수행자 기입 필요",
     "detail": ("[현상] ZCRM052·ZCRM053 결과 ALV 첫 컬럼에 '구독'(구독구분)이 표시된다(값은 공백). 1차(09-14)에서는 표시되지 않았다. 시나리오 검증 5-c(비구독은 미표시) 불일치.\n"
                "[원인 추정] 소스는 비구독 모드에서 SUBS_TYPE 에 NO_OUT='X' 를 주지만, 결과 ALV 레이아웃 키(GV_VARIANT-REPORT = 프로그램+화면+컨테이너)가 4개 T-code 공통이다. "
                "C-05 수정으로 기본 레이아웃에 구독구분이 '표시된 열'로 들어가면서(레이아웃 변경 팝업에서 확인) 비구독 T-code 에서도 그 레이아웃이 NO_OUT 을 덮어쓴다.\n"
                "[영향] 값이 공백이라 업무 영향은 작다(Minor). 다만 일반 화면에 불필요한 컬럼이 보인다.\n"
                "[조치 제안] 구독/비구독 T-code 의 레이아웃 키 분리(GV_VARIANT-HANDLE 또는 REPORT 에 구독 구분 추가), 또는 비구독 모드에서는 필드카탈로그에서 SUBS_TYPE 을 TECH='X' 로 제외."),
     "images": ["S20_r2_5a_2_result.png", "S20_r2_lay_layout_052.png", "S20_r2_7a_2_result.png"]},
]

evidence = [
    ("P-01", "ZCRM_CICSVC 3건 - 꼬리표·클레임유형", "S12_ZCRM_CICSVC_r2_svc_tag.png"),
    ("P-01", "ZCRM_CICSVC 3건 - 구독여부 / 구독무상 개월수 0·0·5", "S12_ZCRM_CICSVC_r2_svc_subs.png"),
    ("P-01", "ZCRM_CICCOST_M 3건 - 유무상여부 Y, 거래일자 2026.09.18", "S12_ZCRM_CICCOST_M_r2_costm_free.png"),
    ("P-01", "ZCRM_CICCOST_SUB 3건 - 표준금액 0", "S12_ZCRM_CICCOST_SUB_r2_sub.png"),
    ("A-01", "ZCRMR194 선택화면", "S20_r2_194_1_selection.png"),
    ("A-01", "실행 결과 - '청구 저장/확정완료', 전체 3, 에러 혹은 유상 3", "S20_r2_194_2_result.png"),
    ("A-02", "ZCRM_REQ - 요청번호 2026091800001 품목 10/20/30, 청구요청확인 X, 꼬리표 9281~9283 (A-03~A-05)", "S12_ZCRM_REQ_r2_req.png"),
    ("A-06", "ZCRM_REQNO - 2026.09.18 / 1", "S12_ZCRM_REQNO_r2_reqno.png"),
    ("A-07", "ZCRM_AUTO_ERRLOG - 3건 '가용수량이 부족합니다' (A-08)", "S12_ZCRM_AUTO_ERRLOG_r2_err.png"),
    ("A-07", "ZCRM_CICCOST_M - 무상 상태 공백(오더 미생성)", "S12_ZCRM_CICCOST_M_r2_costm.png"),
    ("B-01", "ZCRM052 선택화면 - 센터조치 + 미접수 (B-06)", "S20_r2_5a_1_selection.png"),
    ("B-01", "ZCRM052 결과 - TC1 1건 '무상', 제목 (1). 첫 컬럼에 '구독' 이 표시됨(B-03 결함)", "S20_r2_5a_2_result.png"),
    ("B-02", "ZCRM052 TC1 세부사항", "S20_r2_5a_3_detail.png"),
    ("B-03", "ZCRM052 레이아웃 변경 - 구독구분이 '표시된 열' 맨 위에 있음", "S20_r2_lay_layout_052.png"),
    ("B-05", "ZCRM052 서비스점조치 + TC1~TC3 → (0)", "S20_r2_5b_2_result.png"),
    ("C-01", "ZCRM052S 선택화면 - '구독조치', 서비스점조치만", "S20_r2_6a_1_selection.png"),
    ("C-02", "[수정 확인] ZCRM052S 결과 - 제목 (2), 구독 S, 구독무상(N년이내)/(N년이후) (C-03~C-07)", "S20_r2_6a_2_result.png"),
    ("C-03", "ZCRM052S TC2 세부사항", "S20_r2_6a_3_detail_row0.png"),
    ("D-01", "ZCRM053 결과 - '회수품적합성관리 조회 (1)', 저장 비활성", "S20_r2_7a_2_result.png"),
    ("D-02", "[수정 확인] ZCRM053S 선택화면 - 덤프 없이 실행, '[구독전용] … 조회'", "S20_r2_7b_1_selection.png"),
    ("D-02", "[수정 확인] ZCRM053S 결과 - 조회 (2), 구독표기 정상, 저장 비활성", "S20_r2_7b_2_result.png"),
    ("D-02", "ZCRM053S TC3 세부사항 - 구독구분 S, '구독무상(N년이후)'", "S20_r2_7b_3_detail_row1.png"),
    ("E-01", "ZCRM052S TC2 입력", "S20_r2_8a_1_entered.png"),
    ("E-01", "저장 확인 팝업", "S20_r2_8a_2_confirm.png"),
    ("E-01", "저장 완료 - '정상적으로 저장되었습니다.'", "S20_r2_8a_3_after_save.png"),
    ("E-01", "ZCRM_RECPT - 접수번호 2026091800014", "S12_ZCRM_RECPT_r2_recpt.png"),
    ("E-02", "ZCRM052S 접수 조회 - TC2 1건", "S20_r2_8b_2_result.png"),
    ("E-02", "TC2 세부사항 - 접수번호·적합수량", "S20_r2_8b_3_detail.png"),
    ("G-01", "회귀 - ZCRM052 센터조치 미접수 2026.07.01~15 (391)", "S20_r2_g1a_result.png"),
    ("G-01", "회귀 - ZCRM052 서비스점조치 미접수 (619)", "S20_r2_g1c_result.png"),
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
    "cover": {"C17": "2026.09.18"},
    "revision": {"B5": "0.02", "C5": "1차 FAIL 3건(D-02·C-05·C-07) 수정 후 새 데이터로 전체 재테스트", "D5": "2026.09.18", "E5": "※작성자 기입 필요", "F5": "", "G5": ""},
    "prog_header": {"B1": "서비스 - 회수품 관리", "D1": "※개발자명 기입 필요",
                    "B2": "SR03000035400 보일러 서비스 구독 대리점 비용 정산 - 무상 매출처리 자동화(ZCRMR194) → 회수품 적합성 관리(ZCRM052/052S/053/053S) [재테스트]",
                    "D2": "ZCRMR194 / ZCRMR052"},
    "prog_rows": prog_rows,
    "common": {k: {"target": v[0], "result": v[1], "note": v[2]} for k, v in common.items()},
    "fail_rows": fail_rows,
    "evidence": ev_out,
    "evidence_sheet": "테스트증적(캡처)",
    "legend": {"blank": "결과란이 비어 있으면 미실행이다(사유는 테스트 데이터 열). '[재테스트]' 는 1차(v0.01, 09-14) FAIL 항목의 수정 확인 결과. PASS 근거 화면은 '테스트증적(캡처)' 시트에 케이스 ID 별로 넣었다.",
               "note": "KDQ 210, 2026-09-18 재테스트. 테스트 데이터(상담번호 2026091891001~003, 생성자 ZTEST0918B)와 ZCRMR194·052S 가 만든 데이터(요청번호·에러로그·접수번호)는 "
                       "1차와 같이 삭제하지 않고 KDQ 에 남겨 두었다. 상세는 '테스트데이터' 시트."},
}
io.open(OUT, "w", encoding="utf-8").write(json.dumps(payload, ensure_ascii=False, indent=1))
res = [r["result"] for r in prog_rows]
print("prog rows:", len(res), "PASS", res.count("PASS"), "FAIL", res.count("FAIL"), "미실행", res.count(""))
cr = list(payload["common"].values())
print("common Y:", sum(1 for v in cr if v["target"] == "Y"), "PASS", sum(1 for v in cr if v["result"] == "PASS"), "FAIL", sum(1 for v in cr if v["result"] == "FAIL"))
print("evidence:", len(ev_out), "missing:", missing)
