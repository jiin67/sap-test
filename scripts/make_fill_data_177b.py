# -*- coding: utf-8 -*-
# make_fill_data_177b.py : v0.03 payload - D-01 / D-02 수정 후 재테스트 결과를 v0.02 페이로드에 반영.
#   한 파일로 버전관리하므로 기존 판정은 '테스트 데이터' 열에 남기고 최신 결과로 덮어쓴다.
import io, json, os
from PIL import Image as PILImage

SRC = r"C:\sap-test\results\_fill_data_177.json"
OUT = r"C:\sap-test\results\_fill_data_177b.json"
SHOTS = r"C:\sap-test\results\shots"
THUMB = r"C:\sap-test\results\shots\_tpl_177b"
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


d = json.load(io.open(SRC, encoding="utf-8"))
P, F, N = "PASS", "FAIL", ""
R = "[v0.03 재테스트] "

# --------------------------------------------------- 재수행한 케이스의 새 판정
UPD = {
    "TC-04": (P, R + "ZCRM177A 조치일자 2026.09.19 → 일반 3건(2026091990001·002·013). 혼재 상담건 2026091990007 이 빠져 v0.02 의 4건에서 3건으로 줄었다(정상). "
                  "구독 6건은 계속 제외. 구독여부 컬럼이 기본 레이아웃에서 두 번째 열로 바로 보인다"),
    "TC-05": (P, R + "ZCRM177AS 같은 조건 → 구독 7건(003·004·005·006·007·008·012). 혼재 상담건 2026091990007 이 구독 화면에만 나타난다"),
    "TC-06": (P, R + "조치일자 2026.09.17~09.19 → 2026091990009(09.17) 포함, 2026091990010(09.16) 제외. 기존 09-18 테스트분 2건도 정상 표시"),
    "TC-19": (P, R + "구독여부가 기본 레이아웃 '/PT' 의 표시된 열 2번째로 저장되어 별도 조작 없이 보인다(v0.02 에서는 숨긴 열). "
                  "숨긴 열은 6개(판매조직·불만전표여부·계층구조2·센터명·기사코드·기사명)로 줄었고 내부코드 SUBS_TYPE 은 여전히 목록에 없다"),
    "TC-22": (F, R + "중복 노출과 대표값 문제는 해결됐다 - 혼재 상담건 2026091990007 이 구독 화면에만 1행으로 표시되고, "
                  "같은 상담번호에 구독 품목이 여럿일 때(2026091990008 을 12/0 으로 바꿔 확인) SUBSWRTMONTH 가 작은 값(N년이내)이 대표로 선택된다. "
                  "다만 시나리오가 요구한 '결함으로 기록'(혼재 표시)이 동작하지 않는다 - 에러 메시지 열이 공백. FAIL내역 D-04 참조. 확인 후 2026091990008 은 12 로 원복"),
    "TC-36": (P, R + "조치일자 2026.01.01~09.19, 완료건 포함: ZCRM177A 21,574건 / 약 87초(v0.02 21,575건 73초), ZCRM177AS 13건 / 약 29초(v0.02 19초). "
                  "TIME_OUT·덤프 없음. 혼재 판정 LOOP 의 행별 SELECT 때문에 각각 14초·10초 늘었다. FAIL내역 D-05 참조"),
    "TC-41": (P, R + "조치일자 2026.09.19 기준 대상 상담번호 10건에 대해 ZCRM177A 3건 + ZCRM177AS 7건 = 10건. 두 화면 간 중복 없음(v0.02 는 11건)"),
    "X-01": (P, R + "기본 레이아웃 '/PT' 에 구독여부가 추가 저장되어 두 화면 모두 조작 없이 표시된다"),
    "X-02": (F, R + "증상은 해소됐으나 원인은 남아 있다. GV_VARIANT-REPORT = SY-REPID 만 쓰는 구조(Z00CA_ALV_M 654)는 그대로이고 "
                  "GV_VARIANT-HANDLE 은 소스 전체에 없다(검색 0건). 두 T-code 가 계속 같은 레이아웃 키를 공유하므로 한쪽에서 기본 레이아웃을 다시 저장하면 양쪽이 함께 바뀐다. 심각도 하향"),
}
for r in d["prog_rows"]:
    if r["id"] in UPD:
        res, txt = UPD[r["id"]]
        prev = r["result"] or "미실행"
        r["data"] = txt + "\n〔v0.02 결과: %s〕 %s" % (prev, r["data"])
        r["result"] = res

# --------------------------------------------------- 신규 추가 케이스
NEW = [
    ("X-04", "추가", "혼재 상담건 오류 메시지 표시", "에러 메시지 열에 '구독/비구독 품목이 섞인 상담건' 안내가 표시되어야 함", "Major", F,
     R + "2026091990007 의 에러 메시지 열이 공백이다. 소스(F01 152~161)는 SELECT COUNT( DISTINCT subs_type ) 로 혼재를 판정해 gt_data-msg = TEXT-M90 을 넣지만, "
         "텍스트 기호 M90 이 생성되어 있지 않아(SE38 텍스트요소: M01~M12, M16~M19 뒤에 바로 T04) 런타임에 공백이 대입된다. FAIL내역 D-04"),
    ("X-05", "추가", "혼재 판정 추가 후 조회 성능", "혼재 판정 추가로 조회 시간이 업무 허용범위를 벗어나지 않아야 함", "Minor", F,
     R + "혼재 판정 LOOP 가 gt_data 의 모든 행마다 SELECT COUNT( DISTINCT ) 를 수행한다(중복 제거 전이라 품목 단위 행 수만큼). "
         "ZCRM177A 73→87초, ZCRM177AS 19→29초. 덤프·TIME_OUT 은 없으나 긴 기간 조회에서 누적된다. FAIL내역 D-05"),
    ("X-06", "추가", "구독 화면의 조회 범위 변경 확인", "SQL 필터 제거 후에도 구독 화면 결과가 동일해야 함", "Major", P,
     R + "c~subs_type IN gr_subs 를 주석 처리해 두 화면 모두 전건을 읽고 ABAP 에서 거른다. 2026.01.01~09.19 조회 결과는 구독 13건으로 v0.02 와 동일하고 누락·중복 없음"),
]
d["prog_rows"] += [{"id": c[0], "case": "[%s] %s" % (c[1], c[2]), "expect": c[3], "sev": c[4], "result": c[5], "data": c[6]} for c in NEW]

# --------------------------------------------------- 공통테스트케이스
d["common"]["UC-043"] = {"target": "Y", "result": P, "note": "일반 3건 + 구독 7건 = 대상 10건, 중복 없음(v0.03 재테스트)"}
d["common"]["UC-009"] = {"target": "Y", "result": P, "note": "재테스트 전 구간 덤프·행업 없음"}

# --------------------------------------------------- FAIL 내역
FIX01 = ("\n\n[v0.03 조치결과] 해결 - 기본 레이아웃 '/PT' 에 구독여부가 표시된 열 2번째로 저장되어 두 화면 모두 조작 없이 보인다. "
         "다만 코드는 그대로여서(GV_VARIANT-HANDLE 없음) 두 T-code 의 레이아웃 키 공유는 남아 있다. "
         "한쪽에서 기본 레이아웃을 다시 저장하면 양쪽이 함께 바뀌므로, 화면별 분리가 필요하면 HANDLE 분리를 별도 과제로 남긴다.")
FIX02 = ("\n\n[v0.03 조치결과] 해결 - 개선안 3(혼재 건을 오류로 표시)으로 수정됨을 소스로 확인했다. "
         "F01 116 의 SQL 필터 주석 처리, 150 'SORT gt_data BY csltno ASCENDING subs_type DESCENDING subswrtmonth ASCENDING', "
         "152~161 혼재 판정 LOOP, 164 중복 제거, 167~171 모드별 DELETE. "
         "재테스트에서 2026091990007 은 구독 화면에만 1행으로 나오고 합계는 3+7=10 으로 대상과 일치한다. "
         "동점 처리도 의도대로 동작한다(2026091990008 의 품목 월수를 12/0 으로 바꾸니 '구독무상(N년이내)' 이 대표). "
         "남은 문제는 메시지 미표시(D-04)와 성능(D-05)이다.")
for f in d["fail_rows"]:
    if f["id"] == "D-01":
        f["case"] = "[v0.02 FAIL → v0.03 조치] " + f["case"]
        f["detail"] += FIX01
        f["image_paths"] = list(f["image_paths"]) + [p for p in [thumb("S20_R2_lay_r2_col0.png", 620), thumb("S20_R2_layv_r2_variant.png", 620)] if p]
    elif f["id"] == "D-02":
        f["case"] = "[v0.02 FAIL → v0.03 해결] " + f["case"]
        f["detail"] += FIX02
        f["image_paths"] = list(f["image_paths"]) + [p for p in [thumb("S20_R2_A_r2_a.png", 620), thumb("S20_R2_S_r2_s.png", 620)] if p]

d["fail_rows"].append({
    "id": "D-04", "case": "혼재 상담건 오류 메시지가 표시되지 않음 - 텍스트 기호 M90 미생성 (X-04 / TC-22)",
    "tester": "KKR003-1",
    "detail": "혼재 판정 로직은 들어갔으나 화면의 '에러 메시지' 열이 공백이다. ZCRMR177_1F01 158 행이 'gt_data-msg = TEXT-M90.' 인데, "
              "SE38 → 텍스트요소 → 텍스트기호에 M90 이 없다(M01~M12, M16~M19 다음이 바로 T04). "
              "ABAP 은 존재하지 않는 텍스트 기호를 오류 없이 공백으로 처리하므로 활성화는 되지만 런타임에 빈 값이 들어간다.\n"
              "조치안: SE38 → ZCRMR177_1 → 텍스트요소 → 텍스트기호에 M90 을 만들고 '구독/비구독 품목이 섞인 상담건입니다' 를 입력한다. "
              "다른 번호를 쓰기로 했다면 소스의 TEXT-M90 도 함께 바꾼다. 텍스트 기호는 번역 대상이므로 운영 이관 시 텍스트도 함께 이관되는지 확인이 필요하다.",
    "image_paths": [p for p in [thumb("S20_R2_S_r2_s.png", 620), thumb("S24_r177b_m19_M19.png", 620)] if p]})

d["fail_rows"].append({
    "id": "D-05", "case": "혼재 판정이 행마다 DB 를 조회해 조회 시간이 늘어남 (X-05 / TC-36)",
    "tester": "KKR003-1",
    "detail": "ZCRMR177_1F01 152~161 의 LOOP 는 gt_data 의 모든 행에 대해 SELECT COUNT( DISTINCT subs_type ) 을 한 번씩 실행한다. "
              "중복 제거(164 행) 전이라 상담번호가 아니라 품목 단위 행 수만큼 돌고, 화면에서 버릴 행에 대해서도 조회한다. "
              "조치일자 2026.01.01~09.19 기준 ZCRM177A 73→87초, ZCRM177AS 19→29초로 늘었다(덤프·TIME_OUT 은 없음).\n"
              "조치안: 혼재 상담번호를 한 번에 구해 놓고 내부테이블에서 판정한다. 예를 들어 조회 구간의 상담번호별 구독구분 종류 수를\n"
              "  SELECT csltno, COUNT( DISTINCT subs_type ) AS cnt FROM zcrm_cicsvc\n"
              "    WHERE csltno IN @lt_cs AND clmart = 'A' GROUP BY csltno INTO TABLE @DATA(lt_mix)\n"
              "로 한 번에 읽고, LOOP 안에서는 READ TABLE lt_mix 로 판정하면 DB 호출이 1 회로 줄어든다. "
              "중복 제거(164)를 먼저 하고 남은 행에 대해서만 판정하는 것도 방법이다.",
    "image_paths": []})

# --------------------------------------------------- 증적
ev_new = [
    ("TC-04", "[v0.03] ZCRM177A 재테스트 - 일반 3건, 구독여부가 두 번째 열로 기본 표시", "S20_R2_A_r2_a.png"),
    ("TC-05", "[v0.03] ZCRM177AS 재테스트 - 구독 7건, 혼재 상담건 2026091990007 포함", "S20_R2_S_r2_s.png"),
    ("X-01", "[v0.03] 레이아웃 변경 - 구독여부가 표시된 열 2번째, 숨긴 열은 6개", "S20_R2_lay_r2_col0.png"),
    ("X-02", "[v0.03] 레이아웃 선택 - 기본 설정은 여전히 전역 /PT (두 T-code 공유)", "S20_R2_layv_r2_variant.png"),
    ("TC-06", "[v0.03] 조치일자 2026.09.17~19 경계 재확인", "S20_R2_TC06_r2_tc06.png"),
    ("TC-22", "[v0.03] 정렬 동점 처리 - 구독 품목 12/0 → 작은 값(N년이내)이 대표", "S20_R2_tie_r2_tie.png"),
    ("X-04", "[v0.03] SE38 텍스트기호 - M19 다음이 T04, M90 없음", "S24_r177b_m19_M19.png"),
]
for tc, cap, fn in ev_new:
    t = thumb(fn)
    if t:
        d["evidence"].append({"tc": tc, "caption": cap, "path": t})

d["date"] = "2026-09-19"
d["revision"]["B7"] = "0.03"
d["revision"]["C7"] = ("D-01(기본 레이아웃이 구독여부 숨김) / D-02(혼재 상담번호 양쪽 중복) 수정 후 재테스트 - 둘 다 해결. "
                       "신규 D-04(혼재 메시지 미표시, 텍스트기호 M90 미생성) / D-05(혼재 판정 행별 조회로 성능 저하) 확인")
d["revision"]["D7"] = "2026.09.19"
d["revision"]["E7"] = "※작성자 기입 필요"
d["revision"]["F7"] = ""
d["revision"]["G7"] = ""
d["legend"]["blank"] = ("결과란은 최신(v0.03) 기준이다. '[v0.03 재테스트]' 로 시작하는 행은 수정본으로 다시 수행한 것이며 〔v0.02 결과: …〕 에 이전 판정을 남겼다. "
                        "결과란이 비어 있으면 미실행이고 사유는 같은 열에 적었다(※해당없음 / ※미실행(BLOCKED)). X-xx 는 시나리오 외 추가 테스트다.")
d["legend"]["note"] = (d["legend"]["note"] +
                       " v0.03 은 2026-09-19 오후 D-01·D-02 수정본 재테스트분이다. 테스트 데이터는 v0.01~0.02 와 동일한 상담번호 2026091990001~013 을 그대로 사용했고, "
                       "동점 확인을 위해 잠시 바꾼 2026091990008 품목2 의 SUBSWRTMONTH 는 12 로 원복했다.")

io.open(OUT, "w", encoding="utf-8").write(json.dumps(d, ensure_ascii=False, indent=1))
res = [r["result"] for r in d["prog_rows"]]
print("prog rows:", len(res), "PASS", res.count("PASS"), "FAIL", res.count("FAIL"), "미실행", res.count(""))
print("fail rows:", len(d["fail_rows"]), "evidence:", len(d["evidence"]), "missing:", missing)
