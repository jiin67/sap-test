# -*- coding: utf-8 -*-
# make_fill_data_171n.py : v0.03 = full ZCRMR171 run with a SECOND data set (2026-09-18 night), merged into the
#   single version-managed workbook (v0.01 + v0.02 payload from make_fill_data_171m.py).
import io, json, os
from PIL import Image as PILImage

M = json.load(io.open(r"C:\sap-test\results\_fill_data_171m.json", encoding="utf-8"))
OUT = r"C:\sap-test\results\_fill_data_171n.json"
SHOTS = r"C:\sap-test\results\shots"
THUMB = r"C:\sap-test\results\shots\_tpl_171n"
os.makedirs(THUMB, exist_ok=True)
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
A, B, C, D, E, Fd = "100041", "100043", "100044", "100045", "100047", "100050"
R3 = {
    "P-01": (P, "2차 데이터: 상담번호 2026091892001~008(품목 9건), 요청번호 2026091892001, 생성자 ZTEST0918C, 접수번호 2026091800015~023. 구독: 100041·100043(다품목 2건)·100044·100045·100047 / 일반: 100041·100047·100050. 모두 정상 대리점(해지점 아님)"),
    "TC-01": (P, "ZCRM171S 100041 → '1 건 저장 되었습니다.' 노랑, 확정일 2026.09.18, RECPTF (202609, 0000100041, S) 생성"),
    "TC-02": (P, "ZCRM171 100041 → '1 건 저장', 공백 행 추가, S 행 불변"),
    "TC-03": (P, "100043 + 100044 + 100045 → '3 건 저장 되었습니다.' 세 행 노랑, RECPTF S 행 3건(100043 수량 2)"),
    "TC-04": (P, "노랑 100041 → E '이미 확정된 데이터가 포함되어 있습니다.'"),
    "TC-05": (P, "초록 100041(3100069932) → E '이미 확정된 데이터가 포함되어 있습니다.'"),
    "TC-06": (P, "E '라인을 선택 하세요'"),
    "TC-07": (P, "확인 팝업 '아니오' → 빨강 유지, DB 불변"),
    "TC-08": (P, "2026091892006(100047) S→공백 → 본사확정: E '구독 화면에 비구독 상담건이 포함되어 있습니다.' 저장 안 됨"),
    "TC-09": (P, "ZCRM171, 2026091892007(100047) 공백→S → E '일반 화면에 구독 상담건이 포함되어 있습니다.' 저장 안 됨"),
    "TC-10": (N, "※미실행 - 디버거 필요(1차와 동일)"),
    "TC-11": (P, "100043 의 2026091892003 품목 2 만 S→공백 → 본사확정: E '구독 화면에 비구독 상담건이 포함되어 있습니다.' 저장 안 됨(품목 단위 검출 + 메시지 정상)"),
    "TC-12": (P, "100045 노랑 → 확정취소 → 빨강, RECPTF (100045,S) 삭제, 다른 행 불변"),
    "TC-13": (P, "ZCRM171 100050 확정 → 취소 → 빨강"),
    "TC-14": (P, "빨강 100041 → E '확정되지 않았거나 회계전표가 존재하는 데이터가 포함되어 있습니다.'"),
    "TC-15": (P, "초록 100041 → E 동일 메시지, RECPTF·BELNR 유지"),
    "TC-16": (P, "100044 노랑, 2026091892004 S→공백 → 확정취소: E 재검증 메시지, RECPTF·노랑 유지"),
    "TC-17": (P, "100041 S → '전표처리 되었습니다.' 전표 3100069932 / 2026.09.18, 초록, S 행에만 저장"),
    "TC-18": (P, "ZCRM171 100041 → 전표 3100069933, 공백 행에만 저장, S 행 3100069932 그대로"),
    "TC-19": (P, "100043 + 100045 → 3100069934(86,900) / 3100069935(43,450), 각 S 행에 저장"),
    "TC-20": (P, "빨강 100041 → 확인·전기일 후 E '확정 되지 않았거나, 회계전표가 존재하는 …' 전표 없음"),
    "TC-21": (P, "초록 100041 → E 동일 메시지, BELNR 3100069932 유지"),
    "TC-22": (P, "달력 취소 → E '전기일은 필수 입력입니다.'"),
    "TC-23": (P, "100044, 전기일 2020.01.15 → 정보 팝업 + Message 컬럼 '변형 1100과(와) 계정 범주 +의 전기 기간 2020년 001월이 개시되지 않았습니다.' 노랑 유지"),
    "TC-24": (P, "100044 ZTERM E000→E030 → 전표 3100069931, 화면 지급조건 E030·초록. ※전표 라인의 ZTERM 확인(BSEG)은 1차(3100069918)에서 수행"),
    "TC-25": (P, "100043 노랑, 2026091892003 품목1 S→공백 → 전표생성: E 재검증 메시지, 전표 미생성·노랑 유지"),
    "TC-26": (P, "ZCRM171 100050 노랑, 2026091892008 공백→S → 전표생성: E '일반 화면에 구독 상담건이 포함되어 있습니다.' 전표 미생성"),
    "TC-27": (P, "100043 + 100045 선택, 100045 만 S→공백 → 오류, 두 건 모두 전표 미생성"),
    "TC-28": (P, "100041 S 3100069932 → '전표취소 처리 되었습니다.' 노랑, S 행만 공백·CFDATE 유지, 역분개 9930001219"),
    "TC-29": (P, "ZCRM171 100041 3100069933 → 역분개 9930001220, 공백 행만 BELNR 공백"),
    "TC-30": (P, "100043·100044·100045 3건 동시 취소 → 역분개 9930001221 / 9930001222 / 9930001223, 세 행 노랑"),
    "TC-31": (P, "빨강·노랑 100041 → E '확정 되지 않았거나, 회계전표가 존재하지 않는 데이터가 포함되어 있습니다.'"),
    "TC-32": (P, "이미 역분개된 전표를 가진 100036 행 → 정보 팝업 + Message '전표가 이미 역분개되었습니다.' 초록·BELNR 유지"),
    "TC-33": (P, "100043 초록(3100069934), 2026091892003 품목1 S→공백 → 전표취소: E 재검증 메시지, 역분개 없음. 다중(100043+100045)도 미처리"),
    "TC-34": (P, "정상 분기 TC-28·29·30, 오류 분기 TC-32"),
    "TC-35": (P, "100041: S 행과 공백 행이 RECPTF 에 별도 존재"),
    "TC-36": (P, "ZCRM171 100041 확정취소 → 공백 행만 삭제, ZCRM171S 노랑 유지"),
    "TC-37": (P, "ZCRM171S 100041 확정취소 → S 행만 삭제, ZCRM171 노랑 유지"),
    "TC-38": (P, "100041: S=3100069932, 일반=3100069933"),
    "TC-39": (P, "구독 3100069932 취소 후 일반 3100069933 은 RECPTF 유지·ZCRM171 초록 유지"),
    "TC-40": (P, "재실행 시 100041 노랑 / 2026.09.18"),
    "TC-41": (P, "재실행 시 초록 / 3100069932"),
    "TC-42": (P, "재실행 시 노랑, 전표번호 공백"),
    "TC-43": (P, "재실행 시 100041 빨강"),
    "TC-44": (P, "정산: 100050 전표 3100069936 → 전표취소(9930001224) → 확정취소 → 빨강, 구독행 불변. 상세조회 9건 모두 비구독(단, 표시 컬럼 문제는 X-04)"),
    "TC-45": (P, "상세조회 17건 모두 S, 정산조회 구독 대리점만. 제목 '[구독전용] …'. (단, 상세조회 표시 컬럼 문제는 X-04)"),
    "TC-46": (P, "상세조회 판정완료 행(공급업체 123041×1, 123043×2, 123044×1, 123045×1, 123047×1, 각 39,500) = 정산 수량·금액(43,450 / 86,900)"),
    "X-01": (P, "100044 + 100045 선택, 100045 만 S→공백 → 본사확정: 오류, 두 건 모두 저장 안 됨"),
    "X-02": (P, "100044 + 100045 선택, 100045 만 S→공백 → 확정취소: 오류, 두 건 모두 유지"),
    "X-03": (P, "재검증 오류 메시지 정상 표시(TC-08·09·11, X-01)"),
}
hist = {r["id"]: r for r in M["prog_rows"]}
rows = []
for r in M["prog_rows"]:
    r = dict(r)
    res, data = R3[r["id"]]
    prev = "v0.02 " + (r["result"] or "미실행")
    r["result"] = res
    r["data"] = "[v0.03 2차 데이터] " + data + "\n〔이전: " + prev + "〕 " + hist[r["id"]]["data"]
    rows.append(r)
rows.append({"id": "X-04", "case": "[추가(화면)] 상세조회(조회 라디오) 결과 ALV 표시 컬럼", "expect": "상담번호·접수번호·자재 등 상세 컬럼이 모두 표시(1차 v0.01: 26개 컬럼)", "sev": "Major", "result": F,
             "data": "[v0.03 신규] ZCRM171S·ZCRM171 상세조회에서 구독구분·공급업체코드·적합수량·부적합수량·부품가격 5개 컬럼만 보이고 상담번호 등 21개가 숨겨짐 → FAIL내역 참조"})

fails = [dict(f) for f in M["fail_rows"]]
fails.append({"id": "X-04", "case": "[v0.03 신규] 상세조회 결과에 컬럼이 5개만 표시됨", "tester": "※테스트 수행자 기입 필요",
              "detail": ("[현상] ZCRM171S·ZCRM171 의 '조회' 라디오(상세조회) 결과 ALV 에 구독구분·공급업체코드·적합수량·부적합수량·부품가격만 표시된다. "
                         "상담번호·회수품접수번호·자재·서비스점명·판정일 등 21개 컬럼은 레이아웃 변경 팝업의 '숨긴 열'에 들어가 있다. 1차(v0.01)에서는 26개 컬럼이 모두 보였다.\n"
                         "[원인 추정] 조회(상세) ALV 와 정산 ALV 가 같은 레이아웃 키(프로그램+화면+컨테이너)를 쓴다. 정산 화면에서 저장된 기본 레이아웃(컬럼 순서 변경·Message 추가 시점으로 추정)이 "
                         "상세조회에도 적용되면서, 정산 레이아웃에 없는 상세 전용 컬럼이 전부 숨겨졌다. 표시된 5개는 두 ALV 에 공통으로 있는 필드다.\n"
                         "[영향] 상세조회에서 어떤 상담건인지 식별할 수 없다(데이터·건수는 정상).\n"
                         "[조치 제안] (1) 레이아웃 키 분리: GV_VARIANT-HANDLE 을 조회='DTL'/정산='SET' 등으로 구분(구독 T-code 도 분리하면 더 안전). "
                         "(2) 임시로는 정산 화면에서 저장한 기본 레이아웃을 삭제하거나 '기본 설정' 해제."),
              "images": ["S20_N45_n_tc45_detail.png", "S20_N45lay_n_x4_layout_detail.png", "S20_D45_tc45_detail.png"]})
for f in fails:
    if "image_paths" not in f:
        f["image_paths"] = [p for p in (thumb(i, 620) for i in f["images"]) if p]

ev3 = [
    ("P-01", "2차 데이터 - ZCRM_CICSVC 9행(구독여부·대리점)", "S12_ZCRM_CICSVC_d2_svc.png"),
    ("P-01", "ZCRM052 일반 3건 접수 입력", "S20_d2_052g_1_entered.png"),
    ("P-01", "ZCRM052S 구독 6건 접수 저장", "S20_d2_052s_2_saved.png"),
    ("P-01", "ZCRM_RECPT - 접수번호 2026091800015~023", "S12_ZCRM_RECPT_d2_recpt.png"),
    ("P-01", "ZCRM171S 선택화면(정산, 2026.09)", "S20_N00sel_n_sel_new.png"),
    ("P-01", "기준선 ZCRM171S - 새 대리점 5개 빨강", "S20_N00s_n_base_171S.png"),
    ("P-01", "기준선 ZCRM171 - 새 대리점 3개 빨강", "S20_N00a_n_base_171.png"),
    ("TC-06", "선택 없이 본사확정", "S20_N01_n_tc06.png"),
    ("TC-07", "확인 팝업", "S20_N01_n_tc07_popup.png"),
    ("TC-07", "'아니오' 후 빨강 유지", "S20_N01_n_tc07_after.png"),
    ("TC-14", "미확정 건 확정취소 차단", "S20_N01_n_tc14.png"),
    ("TC-20", "미확정 건 전표생성 차단", "S20_N01_n_tc20.png"),
    ("TC-31", "미확정 건 전표취소 차단", "S20_N01_n_tc31_red.png"),
    ("TC-01", "본사확정 전", "S20_N01_n_tc01_before.png"),
    ("TC-01", "본사확정 후 - 100041 노랑", "S20_N01_n_tc01_after.png"),
    ("TC-01", "ZCRM_RECPTF - (100041, S) 생성", "S12_ZCRM_RECPTF_n_tc01.png"),
    ("TC-04", "노랑 건 재확정 차단", "S20_N01_n_tc04.png"),
    ("TC-31", "노랑 건 전표취소 차단", "S20_N01_n_tc31_yellow.png"),
    ("TC-02", "ZCRM171 일반 본사확정 후", "S20_N02_n_tc02_after.png"),
    ("TC-35", "ZCRM_RECPTF - 100041 일반행·구독행 분리", "S12_ZCRM_RECPTF_n_tc35.png"),
    ("TC-36", "ZCRM171 일반 확정취소 후 빨강", "S20_N36_n_tc36_after.png"),
    ("TC-36", "ZCRM_RECPTF - S 행만 남음", "S12_ZCRM_RECPTF_n_tc36.png"),
    ("TC-40", "ZCRM171S 재조회 - 노랑 유지", "S20_N36s_n_tc36_171S_yellow_tc40.png"),
    ("TC-37", "ZCRM171S 구독 확정취소 후 빨강", "S20_N37_n_tc37_after.png"),
    ("TC-37", "ZCRM_RECPTF - 일반행만 남음", "S12_ZCRM_RECPTF_n_tc37.png"),
    ("TC-37", "ZCRM171 은 노랑 유지", "S20_N37a_n_tc37_171_yellow.png"),
    ("TC-43", "확정취소 후 재조회 - 빨강", "S20_N43_n_tc43_requery_red.png"),
    ("TC-11", "SE16N - 2026091892003 품목 2 S→공백", "S21_n11_flip_after.png"),
    ("TC-11", "본사확정 → 재검증 오류, 100043 빨강 유지", "S20_N11_n_tc11_after.png"),
    ("TC-08", "SE16N - 2026091892006 S→공백", "S21_n08_flip_after.png"),
    ("TC-08", "본사확정 → '구독 화면에 비구독 상담건이…'", "S20_N08_n_tc08_after.png"),
    ("TC-09", "ZCRM171 → '일반 화면에 구독 상담건이…'", "S20_N09_n_tc09_after.png"),
    ("X-01", "다중 확정(100045 불일치) 실행 전", "S20_NX1_n_x1_before.png"),
    ("X-01", "오류, 두 건 모두 빨강 유지", "S20_NX1_n_x1_after.png"),
    ("X-01", "ZCRM_RECPTF - 100044·100045 행 없음", "S12_ZCRM_RECPTF_n_x1.png"),
    ("TC-03", "다중 본사확정 - '3 건 저장 되었습니다.'", "S20_N03_n_tc03_after.png"),
    ("TC-03", "ZCRM_RECPTF - S 행 3건", "S12_ZCRM_RECPTF_n_tc03.png"),
    ("X-02", "다중 확정취소(100045 불일치) - 오류, 모두 노랑 유지", "S20_NX2_n_x2_after.png"),
    ("TC-16", "확정취소 직전 변경 - 오류, 100044 노랑 유지", "S20_N16_n_tc16_after.png"),
    ("TC-16", "ZCRM_RECPTF - 행 유지 (X-02 포함)", "S12_ZCRM_RECPTF_n_x2_16.png"),
    ("TC-12", "확정취소 확인 팝업", "S20_N12_n_tc12_popup.png"),
    ("TC-12", "확정취소 후 100045 빨강", "S20_N12_n_tc12_after.png"),
    ("TC-12", "ZCRM_RECPTF - 100045 행 삭제", "S12_ZCRM_RECPTF_n_tc12.png"),
    ("TC-13", "ZCRM171 100050 확정", "S20_N13_n_tc13_confirmed.png"),
    ("TC-13", "ZCRM171 100050 확정취소", "S20_N13_n_tc13_cancelled.png"),
    ("TC-22", "전표생성 확인 팝업", "S20_N22_n_tc22_confirm.png"),
    ("TC-22", "전기일 달력", "S20_N22_n_tc22_calendar.png"),
    ("TC-22", "달력 취소 → '전기일은 필수 입력입니다.'", "S20_N22_n_tc22_after.png"),
    ("TC-23", "마감기간 전기일 → 오류 정보 팝업", "S20_N23_n_tc23_info.png"),
    ("TC-23", "Message 컬럼에 FI 오류 표시, 100044 노랑 유지", "S20_N23_n_tc23_after.png"),
    ("TC-24", "지급조건 E030 입력", "S20_N24_n_tc24_zterm.png"),
    ("TC-24", "전표생성 후 100044 초록 / 3100069931 / E030", "S20_N24_n_tc24_after.png"),
    ("TC-17", "구독 전표생성 전", "S20_N17_n_tc17_before.png"),
    ("TC-17", "구독 전표생성 후 - 100041 초록 / 3100069932", "S20_N17_n_tc17_after.png"),
    ("TC-17", "ZCRM_RECPTF - S 행에만 저장", "S12_ZCRM_RECPTF_n_tc17.png"),
    ("TC-18", "ZCRM171 일반 전표생성 후 - 3100069933", "S20_N18_n_tc18_after.png"),
    ("TC-38", "ZCRM_RECPTF - 일반 3100069933 / 구독 3100069932", "S12_ZCRM_RECPTF_n_tc38.png"),
    ("TC-41", "전표생성 후 재조회 - 초록", "S20_N41_n_tc41_requery_green.png"),
    ("TC-05", "초록 건 본사확정 차단", "S20_N41_n_tc05.png"),
    ("TC-15", "초록 건 확정취소 차단", "S20_N41_n_tc15.png"),
    ("TC-21", "초록 건 전표 재생성 차단", "S20_N41_n_tc21.png"),
    ("TC-25", "SE16N - 2026091892003 품목 1 S→공백", "S21_n25_flip_after.png"),
    ("TC-25", "전표생성 → 재검증 오류, 100043 노랑 유지", "S20_N25_n_tc25_after.png"),
    ("TC-27", "다중 전표생성(100045 불일치) → 오류, 두 건 노랑 유지", "S20_N27_n_tc27_after.png"),
    ("TC-26", "ZCRM171 전표생성 → '일반 화면에 구독 상담건이…' 100050 노랑 유지", "S20_N26_n_tc26_after.png"),
    ("TC-25", "ZCRM_RECPTF - 100043·100045·100050 BELNR 공백", "S12_ZCRM_RECPTF_n_tc25_27.png"),
    ("TC-19", "다중 전표생성 후 - 3100069934 / 3100069935", "S20_N19_n_tc19_after.png"),
    ("TC-19", "ZCRM_RECPTF", "S12_ZCRM_RECPTF_n_tc19.png"),
    ("TC-33", "SE16N - S→공백", "S21_n33_flip_after.png"),
    ("TC-33", "전표취소 → 재검증 오류, 100043 초록 유지", "S20_N33_n_tc33_after.png"),
    ("TC-33", "다중 전표취소도 미처리", "S20_N33_n_tc33_multi_after.png"),
    ("TC-28", "구독 전표취소 확인 팝업", "S20_N28_n_tc28_confirm.png"),
    ("TC-28", "구독 전표취소 후 100041 노랑", "S20_N28_n_tc28_after.png"),
    ("TC-39", "ZCRM_RECPTF - 구독행만 공백, 일반 3100069933 유지", "S12_ZCRM_RECPTF_n_tc28_39.png"),
    ("TC-39", "ZCRM171 은 초록 유지", "S20_N39a_n_tc39_171_green.png"),
    ("TC-42", "전표취소 후 재조회 - 노랑", "S20_N42_n_tc42_requery_yellow.png"),
    ("TC-29", "ZCRM171 일반 전표취소 후", "S20_N29_n_tc29_after.png"),
    ("TC-32", "이미 역분개된 전표 취소 → 정보 팝업", "S20_N32_n_tc32_info.png"),
    ("TC-32", "Message '전표가 이미 역분개되었습니다.' 초록 유지", "S20_N32_n_tc32_after.png"),
    ("TC-30", "3건 동시 전표취소 후 모두 노랑", "S20_N30_n_tc30_after.png"),
    ("TC-30", "ZCRM_RECPTF - BELNR 공백", "S12_ZCRM_RECPTF_n_tc30.png"),
    ("TC-30", "BKPF 3100069931~936 - 각 역분개 전표 연결", "S12_BKPF_n_fi_all.png"),
    ("TC-44", "ZCRM171 100050 전표생성(3100069936)", "S20_N44_n_tc44_1_posted.png"),
    ("TC-44", "전표취소", "S20_N44_n_tc44_2_fi_cancelled.png"),
    ("TC-44", "확정취소 → 빨강", "S20_N44_n_tc44_3_unconfirmed.png"),
    ("TC-45", "ZCRM171S 상세조회 선택화면", "S20_N45_n_tc45_sel.png"),
    ("TC-45", "ZCRM171S 상세조회 - 17건 모두 S (X-04: 컬럼 5개만 표시)", "S20_N45_n_tc45_detail.png"),
    ("X-04", "레이아웃 변경 팝업 - 상담번호 등 21개가 숨긴 열 (결함)", "S20_N45lay_n_x4_layout_detail.png"),
    ("TC-44", "ZCRM171 상세조회 - 9건 모두 비구독 (X-04 동일)", "S20_N44d_n_tc44_detail.png"),
    ("종료", "종료 시점 ZCRM171S", "S20_NFs_n_final_171S.png"),
    ("종료", "종료 시점 ZCRM171", "S20_NFa_n_final_171.png"),
    ("종료", "종료 시점 ZCRM_RECPTF 2026.09", "S12_ZCRM_RECPTF_n_final.png"),
    ("종료", "ZCRM_CICSVC - SUBS_TYPE 원복 확인", "S12_ZCRM_CICSVC_n_final_svc.png"),
]
ev = list(M["evidence"])
for tc, cap, fn in ev3:
    t = thumb(fn)
    if t:
        ev.append({"tc": "[v0.03] " + tc, "caption": cap, "path": t})

payload = dict(M)
payload["prog_rows"] = rows
payload["fail_rows"] = fails
payload["evidence"] = ev
rev = dict(M["revision"])
rev.update({"B7": "0.03", "C7": "다른 테스트 데이터(2차: 대리점 100041·043·044·045·047·050)로 전체 재수행 → 기존 49건 PASS 유지, 신규 FAIL 1건(X-04 상세조회 컬럼 숨김)", "D7": "2026.09.18", "E7": "※작성자 기입 필요", "F7": "", "G7": ""})
payload["revision"] = rev
payload["common"] = dict(M["common"])
payload["common"]["UC-093"] = {"target": "Y", "result": "FAIL", "note": "상세조회 결과에 컬럼이 5개만 표시(X-04, v0.03 신규)"}
payload["legend"] = {
    "blank": "결과란은 최신(v0.03, 2차 데이터 전체 재수행) 기준이다. 테스트 데이터 열 첫머리가 이번 결과이고, 〔이전: …〕 뒤에 v0.02·v0.01 내용을 남겼다. 결과란이 비어 있으면 미실행.",
    "note": "KDQ 210. v0.01 = 09-18 새벽 1차, v0.02 = 09-18 22시 수정분 재테스트, v0.03 = 09-18 23시 2차 데이터 전체 재수행. 증적 시트는 [v0.01] → [v0.02] → [v0.03] 순서. "
            "v0.03 에서 생성한 전표 3100069931~936 은 모두 취소했고, 바꾼 SUBS_TYPE 은 원복했다. 데이터 상세는 '테스트데이터'(1차 데이터)·'테스트데이터_v0.03'(2차 데이터) 시트.",
}
io.open(OUT, "w", encoding="utf-8").write(json.dumps(payload, ensure_ascii=False, indent=1))
res = [r["result"] for r in rows]
print("rows", len(rows), "PASS", res.count("PASS"), "FAIL", res.count("FAIL"), "미실행", res.count(""))
print("fails", len(fails), "evidence", len(ev), "missing", missing)
