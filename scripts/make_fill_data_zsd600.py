# -*- coding: utf-8 -*-
# make_fill_data_zsd600.py : JSON payload for the unit-test result workbook of the
#   Z33SDR600 / ZSD600 improvement re-test (2026-09-14~15, KDQ 210), consumed by fill_template_subs.ps1
import io, json, os
from PIL import Image as PILImage

OUT = r"C:\sap-test\results\_fill_data_zsd600.json"
SHOTS = r"C:\sap-test\results\shots"
THUMB = r"C:\sap-test\results\shots\_tpl_zsd600"
os.makedirs(THUMB, exist_ok=True)
DATE = "2026-09-15"
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
HOLD_SO = "※미실행 - 판매오더 생성(CREA_ORD) 선행 필요. 지정 자재 PNGB3500035L001 ATP 점검 불가로 생성 차단(PRE-01 참조)"
HOLD_DL = "※미실행 - 납품문서 선행 필요(판매오더 생성이 PRE-01 로 차단)"
HOLD_DATA = "※미실행 - 해당 전제조건을 갖춘 QA 주문 지정 필요"
SKIP_R4 = "※테스트 제외 - 2차 개발 미구현(사용자 확인)"
SKIP_LANG = "※테스트 제외 - 사용자 요청(다국어 로그인 미수행)"

# (id, 그룹, 테스트 케이스, 예상결과, 심각도, 결과, 테스트 데이터/실제 결과)
CASES = [
    ("CHK-01", "선행", "수정본 KDQ 반영 확인",
     "이력 테이블 Z33SDT644 존재, 납품삭제 버튼(DEL_DELIVERY) 분리", "Major", P,
     "SE11: Z33SDT644(ID/BSTKD/VERSION/SEQNO/MATNR_OLD/MATNR/VBELN/ERNAM/ERDAT/ERZET) 존재. 헤더 ALV 툴바 tb18=CRE_DELIVERY, tb19=DEL_DELIVERY 로 분리(09-08 에는 id 중복). 인클루드 활성버전 TR KDDK9A01PH"),
    ("PRE-01", "선행", "판매오더 생성(CREA_ORD) - 지정 테스트 데이터",
     "BSTKD 2190 품목을 PNGB3500035L001 로 변경 후 CREA_ORD → 판매오더 생성, Z33SDT644 이력 기록", "Critical", F,
     "BSTKD 2190 / ID 06aa91b1-7b8c-11f1-9bb5-000c298aeffe / SEQNO 1, 자재 ML00007621→PNGB3500035L001(메모리 반영 확인). CREA_ORD 실행 시 상태바 E '품목 PNGB3500035L001에 대해 ATP를 점검할 수 없습니다.' → 판매오더 미생성, Z33SDT644 미기록. "
     "MARA(FERT, EA)·MARC 6100(가용성점검그룹 02)/1100(Z2) 모두 존재하므로 플랜트 미확장 원인 아님 → FAIL내역 참조"),
    # ---------------- R3 ----------------
    ("R3-01", "R3", "정상 납품문서 생성", "납품문서 1건 생성, 성공 로그·납품번호 표시(LIKP/LIPS/ALV-VBELN_VL)", "Critical", N, HOLD_SO),
    ("R3-02", "R3", "판매오더 미생성 주문", "납품 생성 차단, 판매오더 없음 오류 로그", "Major", P,
     "BSTKD 2190(VBELN 공백) 선택 → 납품문서 생성 → 예/아니오 확인창 → 아니오 → VBELN_VL 공백 유지. (예 선택 시 로그 RESULT=E '판매오더가 생성되지 않았습니다.' 확인)"),
    ("R3-03", "R3", "기존 납품문서 중복 생성 방지", "신규 납품 미생성, 기존 납품 존재 메시지", "Major", N, HOLD_DL),
    ("R3-04", "R3", "미완료 판매오더", "생성 차단, 미완료 품목/상태 메시지(VBUV)", "Major", N, HOLD_DATA + "(VBUV 미완료 판매오더)"),
    ("R3-05", "R3", "확정수량 부족", "생성 차단, 확정수량 부족 품목 표시", "Major", N, HOLD_DATA + "(KBMENG<KWMENG)"),
    ("R3-06", "R3", "저장위치 누락", "납품 미생성, BAPI 오류 로그", "Major", N, HOLD_DATA + "(VBAP-LGORT 공백)"),
    ("R3-07", "R3", "납품분할 감지", "납품 수≠1 이면 Rollback, 일부만 남지 않음", "Critical", N, HOLD_DATA + "(분할조건 주문)"),
    ("R3-08", "R3", "저장위치만 다른 품목", "분할조건 없으면 납품 1건, 품목별 LGORT 반영", "Major", N, HOLD_DATA + "(LGORT 1000·5000 혼합)"),
    ("R3-09", "R3", "다건 혼합 처리", "주문별 성공/실패 구분, LOOP 중단 없음", "Major", N, HOLD_DL),
    ("R3-10", "R3", "납품문서번호 재조회", "재실행 후 기존 납품번호 ALV 표시", "Major", N, HOLD_DL),
    # ---------------- DEL ----------------
    ("DEL-01", "DEL", "미출고 납품문서 정상 삭제", "납품 삭제, ALV 납품번호 초기화", "Critical", N, HOLD_DL),
    ("DEL-02", "DEL", "출고전기 완료 납품 삭제", "삭제 차단, 출고전기 완료 메시지", "Critical", N, HOLD_DL + ". 출고전기는 비가역이라 별도 승인 필요"),
    ("DEL-03", "DEL", "분할 납품 삭제", "자동 삭제 차단", "Major", N, HOLD_DL),
    ("DEL-04", "DEL", "삭제 대상 없음", "삭제 대상 없음 메시지", "Minor", P,
     "BSTKD '1422 от 12.05.2026'(VBELN 4293032, 납품 없음) 선택 → 납품문서 삭제 → 확인창 예 → 상태바 '요청을 완료했습니다. 생성로그를 확인하십시오', 결과 로그 RESULT=E '삭제할 납품문서가 없습니다', VBELN 유지·VBELN_VL 공백"),
    ("DEL-05", "DEL", "다건 삭제 혼합 처리", "건별 성공/실패 구분, 성공분만 삭제", "Major", N, HOLD_DL),
    # ---------------- R4 ----------------
    ("R4-01", "R4", "부품 1건 정상 출고", "출고전기 성공", "Major", N, SKIP_R4),
    ("R4-02", "R4", "부품 다건 정상 출고", "3건 성공, 건별 로그", "Major", N, SKIP_R4),
    ("R4-03", "R4", "정상·오류 혼합 출고", "정상 2건 성공, 오류 1건 실패", "Major", N, SKIP_R4),
    ("R4-04", "R4", "보일러 포함", "보일러 제외 또는 시리얼 미입력 차단", "Major", N, SKIP_R4),
    ("R4-05", "R4", "이미 출고 완료된 납품", "중복 출고 차단", "Major", N, SKIP_R4),
    ("R4-06", "R4", "납품문서 없는 주문", "처리 제외, 오류 표시", "Major", N, SKIP_R4),
    # ---------------- R8 ----------------
    ("R8-01", "R8", "출고요청일 기준 ATP 충분", "정상 색상, 판매오더 생성 가능", "Major", N, HOLD_DATA + "(ATP 통과 자재)"),
    ("R8-02", "R8", "출고요청일 기준 ATP 부족", "부족 품목 빨간색(C610), 생성 차단", "Critical", N, HOLD_DATA + "(ATP 부족 자재)"),
    ("R8-03", "R8", "출고요청일 이후 재고 확보", "요청일 기준 부족 처리", "Major", N, HOLD_DATA),
    ("R8-04", "R8", "출고요청일 공백", "SY-DATUM 기준 검사", "Minor", N, HOLD_DATA + "(VDATU 공백 주문)"),
    ("R8-05", "R8", "요청수량과 ATP 동일", "부족 아님, 정상 처리", "Minor", N, HOLD_DATA),
    ("R8-06", "R8", "자재마스터 미등록", "해당 품목 빨간색, 판매오더 생성 차단, 상태바 메시지", "Major", P,
     "BSTKD 2190 품목 ML00007621(1C 코드, MARA 없음) 헤더 더블클릭 → 상태바 '자재 ML00007621에 대한 마스터 레코드가 없습니다.' 화면 캡처(색상 육안 확인)"),
    ("R8-07", "R8", "기본단위 없음", "기본단위 오류, 생성 차단", "Minor", N, HOLD_DATA + "(MEINS 공백 자재)"),
    ("R8-08", "R8", "ATP 조회결과 없음", "ATP 점검 불가 메시지, 빨간색", "Major", P,
     "PNGB3500035L001 로 변경 후 헤더 더블클릭 → 상태바 E 오류, CREA_ORD 시 '품목 PNGB3500035L001에 대해 ATP를 점검할 수 없습니다.' 로 생성 차단(점검 불가 경로 동작). 원인은 PRE-01"),
    ("R8-09", "R8", "ATP 비대상 자재(DIALOGFLAG N)", "ATP 점검 불가 오류 처리", "Minor", N, "※미확정 - R8-08 과 동일 현상이나 DIALOGFLAG 반환값은 미확인(PRE-01 원인 분석과 함께 확인 필요)"),
    ("R8-10", "R8", "동일 자재 합산수량 검사", "합산 6 기준 차단", "Major", N, HOLD_DATA + "(동일 자재 2품목 주문)"),
    ("R8-11", "R8", "일부 품목만 ATP 부족", "부족 품목만 C610", "Minor", N, HOLD_DATA),
    ("R8-12", "R8", "필요수량 0", "ATP 비교 생략, 오류 없음", "Minor", N, HOLD_DATA + "(ZMENG 0 품목)"),
    ("R8-13", "R8", "CHECK_RULE 검증(CO09 대조)", "ZSD600 과 CO09 판정 일치", "Major", N, HOLD_DATA + " + CO09 대조"),
    # ---------------- R9 ----------------
    ("R9-01", "R9", "정상 자재 변경", "자재번호·자재명·기본단위 변경", "Critical", P,
     "BSTKD 2190 / ML00007621 → PNGB3500035L001. 대화상자 Enter 시 자재명 'NGB350-(DELUXE S) 35L,G20,CO,RU,NR-15SW' 해석, 저장(F8) 후 재조회 없이 품목 ALV: MATNR PNGB3500035L001 / MAKTX 표시 / MEINS ******→EA. "
     "※저장은 메모리(GT_Z33SDS601) 반영이며 DB(Z33SDT641/644)는 판매오더 생성 시 기록됨(SE16N 으로 확인)"),
    ("R9-02", "R9", "품목 미선택", "변경 차단, 선택 안내", "Minor", P, "품목 선택 없이 자재 변경 → 대화상자 열리지 않음(차단). 09-08 빌드에서는 '변경할 품목을 한 건만 선택하십시오' 표시 확인"),
    ("R9-03", "R9", "품목 다중 선택", "한 건만 선택 메시지", "Minor", P, "BSTKD 2396(품목 14건) 2건 선택 → 대화상자 열리지 않음(차단)"),
    ("R9-04", "R9", "동일 자재 입력", "변경 차단, 기존 자재 유지", "Minor", P,
     "PNGB3500035L001 로 변경(메모리) 후 같은 자재 재입력·저장 → 자재 유지(PNGB3500035L001). ※별도 안내 메시지(M25) 없이 대화상자만 닫힘 - 메시지 표시가 요건이면 보완 검토"),
    ("R9-05", "R9", "미등록 자재 입력", "자재마스터 오류, 저장 안 됨", "Major", P, "ZZNOEXIST9 입력+Enter → 오류 팝업 '자재 ZZNOEXIST9에 대한 마스터 레코드가 없습니다.', 대화상자 유지, 미저장"),
    ("R9-06", "R9", "플랜트 미등록 자재", "플랜트 미등록 오류", "Major", N,
     HOLD_DATA + "(MARA 존재·해당 플랜트 MARC 없는 자재). ※관찰: 대화상자는 Enter 시 MARA 만 검증(PNGB 는 즉시 통과), 플랜트 검증 시점 확인 필요"),
    ("R9-07", "R9", "판매오더 생성 후 변경", "자재 변경 차단", "Critical", P, "BSTKD '1422 от 12.05.2026'(VBELN 4293032) 품목 → '완료된 판매 문서를 변경할 수 없습니다.', 대화상자 열리지 않음"),
    ("R9-08", "R9", "마감 주문 변경", "자재 변경 차단", "Critical", P, "STATU=C 주문(BSTKD 2345) 품목 → 동일 메시지로 차단"),
    ("R9-09", "R9", "변경 자재 ATP 충분", "변경 유지, 정상 색상", "Major", N, HOLD_DATA + "(ATP 통과 자재)"),
    ("R9-10", "R9", "변경 자재 ATP 부족", "변경 유지, C610, 생성 시 차단", "Major", N, HOLD_DATA + "(ATP 부족 자재). PNGB 는 '점검 불가'라 부족 케이스에 부적합"),
    ("R9-11", "R9", "변경 자재 판매오더 반영", "VBAP-MATNR = 변경 자재", "Critical", N, HOLD_SO),
    ("R9-12", "R9", "자재 변경이력 저장", "Z33SDT644 원본·변경 자재·VBELN 저장", "Critical", N, HOLD_SO),
    ("R9-13", "R9", "변경 없는 주문 이력 미생성", "Z33SDT644 미생성", "Major", N, HOLD_SO),
    ("R9-14", "R9", "여러 품목 변경이력", "SEQNO 별 이력 저장", "Major", N, HOLD_SO),
    # ---------------- HIST ----------------
    ("HIST-01", "HIST", "변경이력 없는 Item 재조회", "Z33SDT641 원본 자재 표시", "Major", P, "BSTKD 2190: Z33SDT644 해당 ID 없음('값이 없습니다') → 품목 원본 ML00007621 표시"),
    ("HIST-02", "HIST", "변경이력 있는 Item 재조회", "Z33SDT644-MATNR 표시", "Critical", N, HOLD_SO),
    ("HIST-03", "HIST", "원본 자재번호 검색", "해당 주문 조회", "Major", N, HOLD_SO),
    ("HIST-04", "HIST", "변경 자재번호 검색", "주문 조회, Item 에 변경자재", "Major", N, HOLD_SO),
    ("HIST-05", "HIST", "관계없는 이력 제외", "완전 키 일치 이력만 적용", "Major", N, HOLD_SO),
    ("HIST-06", "HIST", "MATNR_OLD 불일치 이력 제외", "원본 자재 표시", "Major", N, HOLD_SO),
    ("HIST-07", "HIST", "변경 자재 공백", "원본 자재 유지", "Major", N, HOLD_SO),
    ("HIST-08", "HIST", "최신 변경이력 적용", "ERDAT·ERZET 최신 이력 표시", "Major", N, HOLD_SO),
    ("HIST-09", "HIST", "HASHED TABLE 중복 키 방어", "덤프 없이 최신 이력만 적용", "Critical", N, HOLD_SO),
    ("HIST-10", "HIST", "Header 중복 제거", "동일 ID·BSTKD·VERSION Header 1건", "Major", P,
     "S_MATNR=GXXX002868 조회 → 헤더 38행, ID|BSTKD|VERSION 키 38개 모두 고유(ALV 페이징으로 전행 확인). ※표본 헤더는 해당 자재 1품목씩이라 다품목 중복 케이스는 미검증"),
    ("HIST-11", "HIST", "다건 조회 성능", "지연·덤프 없이 조회", "Major", P, "생성일 2026.01.01~09.08 전체: 조회 2,208건 / 로그 2,232건, 덤프·과도 지연 없음"),
    # ---------------- UI ----------------
    ("UI-01", "UI", "ATP 부족 행 색상", "부족 Item 만 C610", "Minor", N, HOLD_DATA + "(ATP 부족 주문). R8-06/R8-08 캡처로 오류 품목 색상은 확인 가능"),
    ("UI-02", "UI", "ALV Refresh", "재조회 없이 변경 자재·색상 반영", "Minor", P, "R9-01 저장 직후 품목 ALV 에 PNGB3500035L001/자재명/EA 반영(재조회 없음)"),
    ("UI-03", "UI", "한국어 메시지", "한국어 메시지 정상, 잘림 없음", "Minor", P, "모든 오류·차단·로그 메시지 한국어 정상 표시('완료된 판매 문서를 변경할 수 없습니다.', '삭제할 납품문서가 없습니다' 등)"),
    ("UI-04", "UI", "영어 메시지", "EN 로그인 시 영어 표시", "Minor", N, SKIP_LANG),
    ("UI-05", "UI", "러시아어 메시지", "RU 로그인 시 러시아어 표시", "Minor", N, SKIP_LANG),
    ("UI-06", "UI", "다건 처리 로그", "성공·실패 주문별 표시(SALV Popup)", "Minor", N,
     "※부분 확인 - 결과 로그 팝업 구조(BSTKD/VBELN_VA/VBELN_VL/RETCODE/RESULT/RESULT_MESSAGE 행 단위) 확인(DEL-04 캡처). 성공·실패 혼합 다건 실행은 판매오더 선행 필요"),
]
prog_rows = [{"id": c[0], "case": "[%s] %s" % (c[1], c[2]), "expect": c[3], "sev": c[4], "result": c[5], "data": c[6]} for c in CASES]

# ---------------- 공통테스트케이스 ----------------
common = {}


def setc(nums, target, result, note=""):
    for n in nums:
        common["UC-%03d" % n] = (target, result, note)


NA = "해당 없음"
setc(range(1, 120), "N", "", NA)
setc([3], "Y", P, "자재변경 입력 검증: 미등록 자재 오류 팝업(R9-05), 빈 자재 저장 차단")
setc([8], "Y", P, "선택화면 기본값: 생성일 어제~오늘, 주문상태 A~B(조회)/C(동기화), DB연결 CMSRUS_QA")
setc([9, 10, 11], "Y", P, "조회·자재변경·납품/판매오더 버튼 실행 중 덤프·대기·멈춤 없음(ATP 실패도 메시지 처리)")
setc([13], "Y", P, "자재변경 저장은 메모리 반영 설계, 판매오더 생성 시 Z33SDT644 기록(설계대로). 이력 기록 자체는 R9-12 보류")
setc([14], "Y", P, "자재변경 후 품목 ALV 값(MATNR/MAKTX/MEINS)이 입력·해석값과 일치(R9-01)")
setc([16], "Y", "", "판매오더·납품 중복 생성 방지는 생성 선행 필요(R3-03 보류)")
setc([40], "Y", P, "조회조건(1C ID·외부주문번호·자재번호·생성일·상태) 정상")
setc([43], "Y", P, "ID 조회 1건, BSTKD 2190 2건(버전별), 자재 GXXX002868 38건 등 기대치 일치")
setc([45], "Y", P, "자재변경 값이 화면에 즉시 반영(UI-02)")
setc([46], "Y", P, "헤더 더블클릭 → 품목 ALV·ATP 점검, VBELN 더블클릭 VA03 연결 정상")
setc([49, 50, 51], "Y", P, "연 단위 2,208행 조회 덤프·대기·멈춤 없음")
setc([54], "Y", P, "수행옵션 라디오에 따라 상태 필드(S_STAT1/2/3)·DB연결 필드 연동")
setc([55], "Y", P, "ID+BSTKD+자재+생성일+상태 다중조건 동시 적용")
setc([57], "Y", P, "자재 조회 시 헤더 중복 없음(HIST-10)")
setc([58], "Y", P, "헤더 미선택 시 '헤더라인을 선택하십시오', 품목 미선택·다중선택 차단")
setc([59], "Y", "", "납품 삭제 후 화면 Clear 는 납품 생성 선행 필요(DEL-01 보류)")
setc([61], "Y", P, "납품 생성/삭제·판매오더 생성 전 예/아니오 확인창")
setc([62, 63, 64, 65], "Y", P, "삭제 확인·실행 중 덤프·행업 없음(대상 없음 시 로그 메시지)")
setc([66, 67], "Y", "", "실제 삭제 케이스 보류(DEL-01/05)")
setc([87], "Y", P, "개발용 메시지 없음(오류는 사용자 메시지·로그 팝업으로 표시)")
setc([88], "Y", P, "완료/마감 주문 자재변경 차단, 조건 미충족 시 버튼 안내")
setc([90], "Y", P, "자재변경 저장 후 재조회 없이 ALV 반영(UI-02)")
setc([93], "Y", P, "화면 제목·컬럼명·버튼명 한국어 정상(UI-03)")
setc([97], "Y", P, "예/아니오 확인창, 오류 팝업 확인 버튼 정상")
setc([99], "Y", P, "오류 메시지 상황 적합. 단 동일자재 재입력 시 안내 없음(R9-04 비고)")
setc([101], "Y", P, "라디오 초기값 주문조회, 동기화 상태 기본 C")
setc([106], "Y", P, "품목 그리드 자재변경 버튼이 주문 상태(VBELN/마감)에 따라 차단")
setc([111], "Y", P, "사용자 친화 메시지(마스터 레코드 없음, ATP 점검 불가 등)")
setc([114], "Y", "", "R4 부품 출고 일괄처리 미구현(2차 개발) - 제외")
assert len(common) == 119, len(common)

fail_rows = [
    {"id": "PRE-01", "case": "[선행] 판매오더 생성(CREA_ORD) - 지정 자재 ATP 점검 불가", "tester": "김지인",
     "detail": ("[현상] BSTKD 2190(ID 06aa91b1-7b8c-11f1-9bb5-000c298aeffe, SEQNO 1) 품목을 ML00007621→PNGB3500035L001 로 변경(대화상자에서 자재명 정상 해석)한 뒤 CREA_ORD 실행 시 "
                "상태바 E '품목 PNGB3500035L001에 대해 ATP를 점검할 수 없습니다.' 로 판매오더가 생성되지 않는다. 헤더 더블클릭 ATP 재점검에서도 '자재 PNGB3500035L001에 대한 마스터 레코드가 없습니다.' 계열 오류.\n"
                "[확인] SE16N MARA: PNGB3500035L001 = FERT, 기본단위 EA, 삭제 없음. MARC: 플랜트 6100(가용성점검그룹 02)·1100(Z2) 존재. 비교 자재 20010472A(1422 주문, 판매오더 4293032 생성됨)도 6100/02 로 동일 구성 → 플랜트 미확장이 원인이 아님.\n"
                "[조치 제안] ATP 함수 ZCRMF02_RFC_MAT_AVAILABILITY 의 RFC 대상 시스템 및 플랜트 결정 로직 확인. 원인 해결 후 R9-11/12/13/14, R3-01/03/07/08/09/10, DEL-01/03/05, HIST-02~09, R8 정상흐름 재개."),
     "images": ["R8_08_item_pngb_atp.png", "R9_01_after_change.png"]},
]

evidence = [
    ("CHK-01", "BSTKD 2190 조회 - 헤더 2건(버전별), 품목 ML00007621(1C 코드, 자재명 공백)", "D_2190_items.png"),
    ("R9-01", "자재변경 저장 후 품목 ALV - PNGB3500035L001 / 자재명 / 기본단위 EA (재조회 없음, UI-02)", "R9_01_after_change.png"),
    ("R8-06", "1C 코드 품목 헤더 더블클릭 - '자재 ML00007621에 대한 마스터 레코드가 없습니다.'", "R8_06_item_1ccode.png"),
    ("R8-08", "PNGB3500035L001 변경 후 ATP 재점검 - 오류 상태(PRE-01 원인)", "R8_08_item_pngb_atp.png"),
    ("DEL-04", "납품삭제 결과 로그 팝업 - '삭제할 납품문서가 없습니다' (UI-06 로그 구조)", "DEL_04_log_popup.png"),
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
    "cover": {"C17": "2026.09.15"},
    "revision": {"B5": "0.01", "C5": "Z33SDR600(ZSD600) 개선분(납품문서 생성/삭제·자재 유효성/ATP·자재 변경·변경이력 재조회) 재테스트 최초 작성", "D5": "2026.09.15",
                 "E5": "김지인", "F5": "", "G5": ""},
    "prog_header": {
        "B1": "SD - 러시아 1C 주문 I/F",
        "D1": "김지인",
        "B2": "[RU] 1C I/F Sales Order Management - 요구사항 RD-003(납품문서 생성)·RD-008(자재 유효성/ATP)·RD-009(자재 변경·이력) 개선분",
        "D2": "Z33SDR600 / ZSD600",
    },
    "prog_rows": prog_rows,
    "common": {k: {"target": v[0], "result": v[1], "note": v[2]} for k, v in common.items()},
    "fail_rows": fail_rows,
    "evidence": ev_out,
    "evidence_sheet": "테스트증적(캡처)",
    "legend": {
        "blank": "결과란이 비어 있으면 미실행이다(사유는 테스트 데이터 열: PRE-01 판매오더 생성 차단, 전제조건 데이터 미지정, R4 미구현, 다국어 미수행). PASS 건의 근거 화면은 '테스트증적(캡처)' 시트에 케이스 ID 별로 넣었다.",
        "note": "KDQ 210, 2026-09-14~15, SAP GUI Scripting 자동 실행. 확인 대화상자의 '예'는 R9-01 자재변경(메모리 반영, /n 으로 폐기)과 납품이 없음이 확인된 주문의 납품삭제(DEL-04)에만 사용. "
                "판매오더 생성은 ATP 오류로 커밋 전에 중단되어 이번 테스트로 생성·변경된 SAP 문서 및 DB 데이터 없음. 상세: results/ZSD600_retest_summary_20260915.md",
    },
}
io.open(OUT, "w", encoding="utf-8").write(json.dumps(payload, ensure_ascii=False, indent=1))
res = [r["result"] for r in prog_rows]
print("prog rows:", len(res), " PASS:", res.count("PASS"), " FAIL:", res.count("FAIL"), " blank:", res.count(""))
cr = payload["common"].values()
print("common Y:", sum(1 for v in cr if v["target"] == "Y"), " PASS:", sum(1 for v in cr if v["result"] == "PASS"), " FAIL:", sum(1 for v in cr if v["result"] == "FAIL"))
print("evidence:", len(ev_out), " missing:", missing)
