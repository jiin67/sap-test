# -*- coding: utf-8 -*-
# make_fill_data_171b.py : payload for the 2026-09-18 (evening) RETEST of ZCRMR171 after the developer fix (v0.02).
#   Takes the case catalog / common checklist of make_fill_data_171.py (v0.01) and patches the retested rows.
import io, json, os
from PIL import Image as PILImage

OUT = r"C:\sap-test\results\_fill_data_171b.json"
SHOTS = r"C:\sap-test\results\shots"
THUMB = r"C:\sap-test\results\shots\_tpl_171b"
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
src = io.open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "make_fill_data_171.py"), encoding="utf-8").read()
exec(src[src.index("MSG = "):src.index("prog_rows = [")])          # MSG, CASES
exec(src[src.index("common = {}"):src.index("T = \"※")])            # common checklist (v0.01)

RE = "[재테스트 09-18 저녁, 수정본] "
KEEP = "[1차 결과 유지 - 이번 수정과 무관, 재수행 안 함] "
NEW = {
    "TC-03": (P, RE + "100037 + 100038 선택 → '2 건 저장 되었습니다.' 두 행 노랑, RECPTF S 행 2건 생성"),
    "TC-08": (P, RE + "2026091890010(100037) S→공백 → 본사확정: E '구독 화면에 비구독 상담건이 포함되어 있습니다.' 저장 안 됨·빨강 유지. 1차의 '0 건 저장' 메시지 문제 해소"),
    "TC-09": (P, RE + "ZCRM171, 2026091890011(100037) 공백→S → 본사확정: E '일반 화면에 구독 상담건이 포함되어 있습니다.' 저장 안 됨·빨강 유지"),
    "TC-12": (P, RE + "100038 노랑 → 확정취소 → 빨강, RECPTF (100038,S) 삭제, 다른 행 불변"),
    "TC-16": (P, RE + "100038 노랑, 2026091890012 S→공백 → 확정취소: E '구독 화면에 비구독 상담건이 포함되어 있습니다.' RECPTF·노랑 유지"),
    "TC-19": (P, RE + "100035 + 100037 → '전표처리 되었습니다.' 3100069928 / 3100069929, 각 S 행에 저장"),
    "TC-23": (P, RE + "1차 FAIL → 수정 확인. 100035, 전기일 2020.01.15 → 정보 팝업 후 ALV 'Message' 컬럼에 "
                    "'변형 1100과(와) 계정 범주 +의 전기 기간 2020년 001월이 개시되지 않았습니다.' 표시. 노랑 유지·BELNR 미저장"),
    "TC-25": (P, RE + "1차 FAIL → 수정 확인. 100037 노랑, 2026091890010 S→공백 → 전표생성: E '구독 화면에 비구독 상담건이 포함되어 있습니다.' 전표 미생성, BELNR 공백, 노랑 유지"),
    "TC-26": (P, RE + "1차 FAIL → 수정 확인. ZCRM171 100034 확정 후 2026091890004 공백→S → 전표생성: E '일반 화면에 구독 상담건이 포함되어 있습니다.' 전표 미생성, 노랑 유지"),
    "TC-27": (P, RE + "1차 FAIL → 수정 확인. 100035 + 100037 선택, 100037 만 S→공백 → 전표생성: 오류 메시지, 두 건 모두 전표 미생성·노랑 유지"),
    "TC-30": (P, RE + "100035(3100069928→9930001216) + 100037(3100069929→9930001217) 동시 취소, 두 행 노랑·BELNR 공백"),
    "TC-32": (P, RE + "이미 역분개된 전표를 가진 100036 행 → 전표취소: 정보 팝업 후 Message 컬럼에 '전표가 이미 역분개되었습니다.' 표시, 초록·BELNR 유지, RECPTF 불변"),
    "TC-33": (P, RE + "1차 FAIL → 수정 확인. 100037 초록(3100069929), 2026091890010 S→공백 → 전표취소: E 메시지, 역분개 없음(BKPF STBLG 공백), 초록·BELNR 유지. "
                    "다중 선택(100035+100037)도 두 건 모두 취소되지 않음"),
    "TC-34": (P, RE + "정상 분기: TC-30. 오류 분기(E_RESULT='N'): TC-32 에서 BELNR·초록 유지 + Message 표시"),
    "TC-44": (P, RE + "ZCRM171 100034: 전표 3100069930 → 전표취소 → 확정취소 → 빨강. 구독 RECPTF 행 불변 (상세조회는 1차 결과 유지)"),
    "X-01": (P, RE + "1차 FAIL → 수정 확인. 100037 + 100038 선택, 100038(2026091890012)만 S→공백 → 본사확정: 오류 메시지, 두 건 모두 저장 안 됨(빨강 유지), RECPTF 변화 없음"),
    "X-02": (P, RE + "1차 FAIL → 수정 확인. 100035 + 100037 선택, 100037 만 S→공백 → 확정취소: 오류 메시지, 두 건 모두 삭제되지 않음(노랑 유지)"),
    "X-03": (P, RE + "1차 FAIL → 수정 확인. 본사확정 재검증 실패 시 '0 건 저장' 대신 재검증 오류 메시지가 표시됨(TC-08·09, X-01)"),
}
rows = []
for cid, grp, name, expect, sev, res, data in CASES:
    if cid in NEW:
        res, data = NEW[cid]
    elif cid == "TC-10":
        pass
    elif cid == "TC-11":
        data = KEEP + data.replace(MSG, "※메시지 덮임 문제는 같은 코드(SET_KUNCR)라 TC-08 재테스트로 해소 확인. 100032 가 전표생성 상태라 이 케이스 자체는 재수행하지 못함")
    elif cid != "P-01":
        data = KEEP + data
    rows.append({"id": cid, "case": "[%s] %s" % (grp, name), "expect": expect, "sev": sev, "result": res, "data": data})
prog_rows = rows

common["UC-003"] = ("Y", P, "구독구분 재검증이 확정·확정취소·전표생성·전표취소 모두에서 처리 전 차단(수정 확인)")
common["UC-066"] = ("Y", P, "다중 확정취소에서 오류 건이 있으면 전체 미처리(X-02 수정 확인)")
common["UC-099"] = ("Y", P, "재검증 오류 메시지 정상 표시, FI 오류는 Message 컬럼에 표시(수정 확인)")
assert len(common) == 119

fail_rows = [
    {"id": "(없음)", "case": "재테스트 범위에서 FAIL 없음", "tester": "※테스트 수행자 기입 필요",
     "detail": ("1차(v0.01) FAIL 8건(TC-23·25·26·27·33, X-01·02·03)을 모두 재테스트해 수정을 확인했다.\n"
                "남은 정리 사항(결함 아님): ZCRM_RECPTF (202609 / 0000100036 / S) 에 이미 역분개된 전표번호 3100069924 가 남아 있어 화면에 초록으로 보인다. "
                "1차 TC-32 오류조건(FB08 수동 역분개) 때문이며, 테이블이 SE16N 편집 불가라 개발자 정리(BELNR·BUDAT 공백)가 필요하다. 아래는 수정 확인 대표 화면."),
     "images": ["S20_R25_r_tc25_after.png", "S20_R23_r_tc23_after.png"]},
]

ev = [
    ("기준선", "재테스트 시작 시점 ZCRM171S (Message 컬럼 추가됨)", "S20_R0s_r_00_baseline_171S.png"),
    ("기준선", "재테스트 시작 시점 ZCRM_RECPTF 2026.09", "S12_ZCRM_RECPTF_r_base.png"),
    ("TC-08", "SE16N - 2026091890010 S→공백", "S21_r08_flip_after.png"),
    ("TC-08", "[수정 확인] 본사확정 → '구독 화면에 비구독 상담건이 포함되어 있습니다.' 100037 빨강 유지 (X-03)", "S20_R08_r_tc08_after.png"),
    ("TC-09", "[수정 확인] ZCRM171 → '일반 화면에 구독 상담건이 포함되어 있습니다.'", "S20_R09_r_tc09_after.png"),
    ("TC-08", "ZCRM_RECPTF - 100037 행 없음", "S12_ZCRM_RECPTF_r_tc0809.png"),
    ("X-01", "다중 선택(100037·100038), 100038 만 불일치 - 실행 전", "S20_RX1_r_x1_before.png"),
    ("X-01", "[수정 확인] 오류 메시지, 두 건 모두 빨강 유지", "S20_RX1_r_x1_after.png"),
    ("X-01", "ZCRM_RECPTF - 두 대리점 모두 행 없음", "S12_ZCRM_RECPTF_r_x1.png"),
    ("TC-03", "정상 다중 본사확정 - '2 건 저장 되었습니다.'", "S20_RT03_r_tc03_after.png"),
    ("TC-03", "ZCRM_RECPTF - 100037·100038 S 행 생성", "S12_ZCRM_RECPTF_r_tc03.png"),
    ("X-02", "[수정 확인] 다중 확정취소(100037 불일치) - 오류 메시지, 100035·100037 모두 노랑 유지", "S20_RX2_r_x2_after.png"),
    ("X-02", "ZCRM_RECPTF - 두 행 모두 유지", "S12_ZCRM_RECPTF_r_x2.png"),
    ("TC-16", "확정취소 직전 변경 - 오류, 100038 노랑 유지", "S20_R16_r_tc16_after.png"),
    ("TC-25", "SE16N - 2026091890010 S→공백", "S21_r25_flip_after.png"),
    ("TC-25", "[수정 확인] 전표생성 → 오류 메시지, 100037 노랑 유지·전표번호 없음", "S20_R25_r_tc25_after.png"),
    ("TC-27", "다중 전표생성(100035·100037, 100037 불일치) - 실행 전", "S20_R27_r_tc27_before.png"),
    ("TC-27", "[수정 확인] 오류 메시지, 두 건 모두 노랑 유지", "S20_R27_r_tc27_after.png"),
    ("TC-25", "ZCRM_RECPTF - 100035·100037 BELNR 공백 (TC-27)", "S12_ZCRM_RECPTF_r_tc25_27.png"),
    ("TC-26", "ZCRM171 100034 본사확정", "S20_R26_r_tc26_confirmed.png"),
    ("TC-26", "[수정 확인] 공백→S 변경 후 전표생성 → '일반 화면에 구독 상담건이 포함되어 있습니다.' 노랑 유지", "S20_R26_r_tc26_after.png"),
    ("TC-26", "ZCRM_RECPTF - 100034 BELNR 공백", "S12_ZCRM_RECPTF_r_tc26.png"),
    ("TC-23", "마감기간 전기일 → 오류 정보 팝업", "S20_R23_r_tc23_info.png"),
    ("TC-23", "[수정 확인] ALV Message 컬럼에 FI 오류 내용 표시, 100035 노랑 유지", "S20_R23_r_tc23_after.png"),
    ("TC-23", "행 세부사항 - Message", "S20_R23_r_tc23_detail_b.png"),
    ("TC-19", "정상 다중 전표생성 - 100035·100037 초록", "S20_RT19_r_tc19_after.png"),
    ("TC-19", "ZCRM_RECPTF - 3100069928 / 3100069929", "S12_ZCRM_RECPTF_r_tc19.png"),
    ("TC-33", "SE16N - 2026091890010 S→공백", "S21_r33_flip_after.png"),
    ("TC-33", "[수정 확인] 전표취소 → 오류 메시지, 100037 초록·전표번호 유지", "S20_R33_r_tc33_after.png"),
    ("TC-33", "[수정 확인] 다중 전표취소(100035·100037)도 두 건 모두 유지", "S20_R33_r_x3_multi_fidel_after.png"),
    ("TC-33", "BKPF 3100069928·929 - 역분개(STBLG) 없음", "S12_BKPF_r_fi_tc33.png"),
    ("TC-32", "이미 역분개된 전표(100036) 전표취소 → 정보 팝업", "S20_R32_r_tc32_info.png"),
    ("TC-32", "Message 컬럼 '전표가 이미 역분개되었습니다.' 초록·전표번호 유지", "S20_R32_r_tc32_after.png"),
    ("TC-30", "정상 다중 전표취소 - 100035·100037 노랑", "S20_RT30_r_tc30_after.png"),
    ("TC-30", "BKPF - 역분개 9930001216 / 9930001217", "S12_BKPF_r_fi_tc30.png"),
    ("TC-12", "정상 확정취소 - 100038 빨강", "S20_RT30_r_tc12_after.png"),
    ("TC-12", "ZCRM_RECPTF - 100038 행 삭제, 나머지 유지", "S12_ZCRM_RECPTF_r_tc30_12.png"),
    ("TC-44", "ZCRM171 100034 전표생성(3100069930)", "S20_RT44_r_tc44_1_posted.png"),
    ("TC-44", "ZCRM171 100034 전표취소", "S20_RT44_r_tc44_2_fi_cancelled.png"),
    ("TC-44", "ZCRM171 100034 확정취소 → 빨강", "S20_RT44_r_tc44_3_unconfirmed.png"),
    ("종료", "종료 시점 ZCRM171S", "S20_RFs_r_final_171S.png"),
    ("종료", "종료 시점 ZCRM_RECPTF 2026.09", "S12_ZCRM_RECPTF_r_final.png"),
    ("종료", "ZCRM_CICSVC - SUBS_TYPE 원복 확인", "S12_ZCRM_CICSVC_r_final_svc.png"),
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
    "cover": {"C17": "2026.09.18"},
    "revision": {"B5": "0.02", "C5": "1차 FAIL 8건(재검증 차단·부분처리·메시지) 수정 후 재테스트", "D5": "2026.09.18", "E5": "※작성자 기입 필요", "F5": "", "G5": ""},
    "prog_header": {"B1": "서비스 - AS 무상 회수품 지원금", "D1": "※개발자명 기입 필요",
                    "B2": "SR03000035400 ZCRMR171 일시불구독 개선 - 본사확정 및 FI 처리 (ZCRM171 / ZCRM171S) [재테스트]", "D2": "ZCRMR171"},
    "prog_rows": prog_rows,
    "common": {k: {"target": v[0], "result": v[1], "note": v[2]} for k, v in common.items()},
    "fail_rows": fail_rows,
    "evidence": ev_out,
    "evidence_sheet": "테스트증적(캡처)",
    "legend": {"blank": "결과란이 비어 있으면 미실행이다. '[재테스트 …]' 로 시작하는 행은 수정본으로 다시 수행한 결과, '[1차 결과 유지 …]' 행은 이번 수정과 무관해 v0.01(09-18 새벽) 결과를 그대로 옮긴 것이다.",
               "note": "KDQ 210, 2026-09-18 22시 재테스트. 1차와 같은 테스트 데이터(기간월 2026.09)를 사용했고, 바꾼 SUBS_TYPE 은 모두 원복했다. 이번에 생성한 전표 3100069928~930 은 모두 취소했다. "
                       "이 시트의 증적은 재테스트분만 담았다. 1차 증적은 v0.01 참조."},
}
io.open(OUT, "w", encoding="utf-8").write(json.dumps(payload, ensure_ascii=False, indent=1))
res = [r["result"] for r in prog_rows]
print("prog rows:", len(res), "PASS", res.count("PASS"), "FAIL", res.count("FAIL"), "미실행", res.count(""))
print("retested:", len(NEW), "evidence:", len(ev_out), "missing:", missing)
