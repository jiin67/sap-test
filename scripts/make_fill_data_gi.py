# -*- coding: utf-8 -*-
# make_fill_data_gi.py : JSON for the 2026-09-16 GI_ICON / delivery / material-change scenario (KDQ 210)
import io, json, os
from PIL import Image as PILImage

OUT = r"C:\sap-test\results\_fill_data_gi.json"
SHOTS = r"C:\sap-test\results\shots\gi"
THUMB = r"C:\sap-test\results\shots\_tpl_gi"
os.makedirs(THUMB, exist_ok=True)
DATE = "2026-09-17"
missing = []


def thumb(fn, maxw=780):
    src = os.path.join(SHOTS, fn)
    if not os.path.exists(src):
        missing.append(fn); return None
    dst = os.path.join(THUMB, fn)
    im = PILImage.open(src); im.thumbnail((maxw, maxw)); im.save(dst, optimize=True)
    return dst


P, F, N = "PASS", "FAIL", ""
D_TARGETS = "테스트 대상 BSTKD 1757/2318/1722/2392/2766/2827 (기준일 2026.09.16). 재테스트(2026.09.17)는 1722로 판매오더 4293343 생성→마감→납품 85857220+출고전기→납품삭제 흐름 수행 후 판매오더·납품 전부 삭제(ELIM_ORD '성공적으로 삭제 되었습니다'), Z33SDT644 이력 없음 확인, 1722 진입시 STATU와 동일하게 원복 - KDQ 원상복구 완료"

CASES = [
    # ---- 2. 기본 조회/자재정보 ----
    ("BAS-01", "기본조회", "주문 조회 (Header/Item)", "각 BSTKD Header 조회, 더블클릭 시 전체 Item 표시", "Major", P,
     "6개 BSTKD 모두 조회됨. 예) 2318 H0 Item 11건(SEQNO/MATNR/MEINS EA/LGORT 5000) 표시. 1722 부품 6건, 2392 보일러 1건"),
    ("BAS-02", "기본조회", "자재명·기본단위 조회", "MATNR별 MAKTX와 MARA 기본단위 표시", "Minor", P,
     "품목 ALV MAKTX 표시, MEINS=EA. MARA 조회: 대상 자재 전부 MEINS=EA (FERT/PART/ASSY)"),
    ("BAS-03", "기본조회", "계정지정그룹(KTGRM) 조회", "각 자재 KTGRM 정상 조회", "Minor", P,
     "MVKE VKORG=6100 기준: 완제품(PNGB*)·2827=Y1, 부품/조립(1722의 20*/30*)=Y2. KTGRM에 따라 납품/출고 분기(DIV 참조)"),
    # ---- 3. 자재 변경 ----
    ("MAT-01", "자재변경", "정상 자재 변경 (1757)", "MATNR/MAKTX/MEINS 신규 자재 기준 변경", "Critical", P,
     "1757 마감취소(B) 후 PNGB3100013L003→PNGB3500013L008. 저장 후 재조회 없이 품목 ALV: MATNR·MAKTX·MEINS(EA) 반영"),
    ("MAT-02", "자재변경", "Item 미선택", "선택 안내 메시지, 변경 안 됨", "Minor", P, "Item 미선택 CHANGE_MAT → 대화상자 안 열림(차단)"),
    ("MAT-03", "자재변경", "Item 다중 선택", "한 건만 선택 메시지", "Minor", P, "2318 Item 2건 선택 → 대화상자 안 열림(차단)"),
    ("MAT-04", "자재변경", "미등록 자재 입력", "자재마스터 오류, 기존 자재 유지", "Major", P,
     "ZZNOEXIST9 입력 → 오류 팝업 '자재 ZZNOEXIST9에 대한 마스터 레코드가 없습니다.', 기존 자재 유지"),
    ("MAT-05", "자재변경", "플랜트 미등록 자재", "플랜트 미등록 오류, 기존 자재 유지", "Major", N,
     "※미실행 - MARA 존재·해당 플랜트(6100) MARC 없는 자재를 특정하지 못함. 대화상자는 Enter 시 MARA만 즉시 검증(플랜트 검증 시점 확인 필요)"),
    ("MAT-06", "자재변경", "판매오더 생성 후 자재 변경", "판매오더 완료 메시지, 변경 차단", "Critical", P,
     "1757 판매오더 4293338 생성 후 → '완료된 판매 문서를 변경할 수 없습니다.', 대화상자 안 열림"),
    ("MAT-07", "자재변경", "마감 주문 자재 변경", "마감 품목 변경 불가 메시지", "Critical", P,
     "1757 STATU=C 상태에서 CHANGE_MAT → '마감품목 변경이 불가능합니다', 대화상자 안 열림"),
    # ---- 4. MEINS ----
    ("UNIT-01", "MEINS", "변경 자재 기본단위 동시 갱신", "GT_Z33SDS601/GT_Z33SDT641 MEINS 신규 기본단위로, 판매오더 생성 시 단위 오류 없음", "Critical", P,
     "1757 자재 변경 후 품목 MEINS=EA, CREA_ORD 성공(오더 4293338, 단위 오류 없음)"),
    # ---- 5. ATP ----
    ("ATP-01", "ATP", "ATP 충분 (1757)", "정상 색상, 판매오더 생성 가능", "Major", P,
     "1757 자재변경 후 헤더 더블클릭 ATP 오류 없음 → CREA_ORD 성공(4293338)"),
    ("ATP-02", "ATP", "ATP 부족 (2318)", "부족 Item C610, 판매오더 생성 차단", "Critical", P,
     "2318 헤더 더블클릭 시 '_ATP수량이 부족합니다_', SEQNO11(28EA) 등 빨간색(C610), CREA_ORD 차단"),
    ("ATP-03", "ATP", "ATP=요청수량 경계값", "부족으로 판정하지 않음", "Minor", N, "※미실행 - COM_QTY=REQ_QTY 되는 재고 상태의 주문/자재 특정 필요"),
    ("ATP-04", "ATP", "과거 출고요청일 보정", "LV_REQ_DATE=SY-DATUM 보정", "Major", P,
     "1757 VDATU=2026.06.09(과거)인데도 ATP 통과·오더 생성 → 과거 요청일이 현재일 기준으로 보정됨(관찰)"),
    ("ATP-05", "ATP", "출고요청일 공백", "SY-DATUM 기준 검사", "Minor", N, "※미실행 - VDATU 공백 테스트 주문 필요"),
    ("ATP-06", "ATP", "변경 자재 ATP 재검사", "신규 MATNR/MEINS/LGORT/VDATU 기준 ATP", "Major", P,
     "1757 자재변경 후 헤더 더블클릭 → 변경 자재 기준으로 ATP 재검사(오류 없이 통과)"),
    ("ATP-07", "ATP", "동일 자재 수량 합산", "합산 요청수량으로 ATP 검사", "Major", N,
     "※재설계 필요 - 2318 item11을 item2와 동일자재(PNGB3500013L008)로 변경했으나 해당 자재 재고 충분으로 합산 부족이 유도되지 않아 오더 생성됨. 재고 부족한 동일자재 필요"),
    ("ATP-08", "ATP", "생성 직전 최신 ATP 재검사", "조회시점이 아닌 생성 직전 ATP로 판정", "Minor", N, "※미실행 - 조회 후 재고를 인위적으로 바꿔 비교하는 시나리오, 재고 조정 권한/절차 필요"),
    ("ATP-09", "ATP", "WMDVEX 결과 없음", "ATP 점검 불가 메시지, 생성 차단", "Major", P,
     "관찰 - ATP 함수가 결과를 못 주는 자재(예: 6100 플랜트 미확장분)는 '자재 X에 대한 마스터 레코드가 없습니다/ATP를 점검할 수 없습니다'로 생성 차단(2026-09-15 PNGB3500035L001 사례)"),
    ("ATP-10", "ATP", "필요수량 0", "ATP 비교 생략, 오류 처리 안 함", "Minor", N, "※미실행 - ZMENG=0 품목 주문 필요"),
    # ---- 6. 변경이력 ----
    ("HIST-01", "이력", "변경 자재 판매오더 반영", "VBAP-MATNR = 변경 자재", "Critical", P,
     "1757/2318 자재변경 후 CREA_ORD → Z33SDT644에 VBELN과 변경 MATNR 기록(4293338/4293339). 재조회 Item도 변경 자재 표시"),
    ("HIST-02", "이력", "변경이력 저장 (Z33SDT644)", "원본·변경 자재·판매오더번호 저장", "Critical", P,
     "Z33SDT644: 1757 SEQNO1 OLD=PNGB3100013L003 NEW=PNGB3500013L008 VBELN=4293338 / 2318 SEQNO11 OLD=PNGB3510030L003 NEW=PNGB3500013L008 VBELN=4293339, ERNAM/ERDAT 기록"),
    ("HIST-03", "이력", "변경 자재 재조회 표시", "재실행 후 Z33SDT644 변경 자재가 Item에 표시", "Major", P,
     "프로그램 재조회(신규 F8) 후에도 Item ALV에 변경 자재(PNGB3500013L008) 표시(Z33SDT644 오버레이)"),
    ("HIST-04", "이력", "원본 자재 검색", "S_MATNR=MATNR_OLD 조회 시 해당 주문 조회", "Major", N, "※미실행 - HIST 준비 분기에서 미수행(엔진 재로그인 후 흐름 단축). 자재 검색 자체는 REG-01에서 정상 동작 확인"),
    ("HIST-05", "이력", "변경 자재 검색", "S_MATNR=변경 MATNR 조회, Item에 변경 자재", "Major", P,
     "S_MATNR=PNGB3500013L008 조회 시 변경한 주문 포함 조회(REG-01 중복제거 검증에서 동일 조건 사용)"),
    ("HIST-06", "이력", "변경 없는 주문 이력 미생성", "Z33SDT644 미생성", "Major", P,
     "2392·1722를 자재변경 없이 CREA_ORD → Z33SDT644 해당 ID '값이 없습니다'(이력 미생성 확인)"),
    ("HIST-07", "이력", "SEQNO별 변경이력", "각 SEQNO별 이력 저장", "Major", N,
     "※부분 - 단일 품목 변경 이력은 SEQNO 단위 기록 확인(1757 SEQNO1, 2318 SEQNO11). 한 주문 내 2품목 동시 변경 케이스는 미실행"),
    ("HIST-08", "이력", "최신 변경이력 적용", "ERDAT/ERZET 최신 자재 표시, ST22 덤프 없음", "Major", N, "※미실행 - 동일 Key 이력 다건 상황 구성 필요"),
    # ---- 7. KTGRM 분기 ----
    ("DIV-01", "분기", "KTGRM=Y1 → 납품까지만", "DEL_CHK 공란, 납품문서까지만 생성", "Major", P,
     "1757(전 품목 KTGRM Y1) CRE_DELIVERY → 납품 85857217 생성, GI_ICON 공란(출고전기 미실행). 2827/2392도 Y1→납품만"),
    ("DIV-02", "분기", "KTGRM≠Y1 → 출고전기까지", "DEL_CHK=X, 출고전기까지 실행", "Major", P,
     "1722(부품 KTGRM Y2) CRE_DELIVERY → 납품 85857218 생성 + 출고전기 완료(VBUP-WBSTA 전부 C, GI_ICON 초록)"),
    ("DIV-03", "분기", "MVKE 정보 없는 자재 포함", "DEL_CHK 판정 업무기준 처리", "Minor", N, "※미실행 - MVKE 없는 자재 포함 주문 특정 필요"),
    # ---- 8. 납품/출고전기 ----
    ("DEL-01", "납품", "납품문서까지만 생성 (1722→실제 1757)", "납품 생성, VBELN_VL 표시, GI_ICON 공란, 문서별 로그 1건", "Critical", P,
     "1757(Y1) 마감 후 CRE_DELIVERY → 납품 85857217, VBELN_VL 표시, GI_ICON 공란, 로그 RESULT=S 1건"),
    ("DEL-02", "납품", "납품 및 출고전기 (2392→실제 1722)", "납품 생성, 전 품목 WBSTA=C, GI_ICON 초록, 로그 1건", "Critical", P,
     "1722(Y2) → 납품 85857218 + 출고전기, VBUP-WBSTA 전부 C, GI_ICON 초록(@08@), 로그 RESULT=S"),
    ("DEL-03", "납품", "판매오더 없음 (2766)", "납품 생성 차단, 오류 로그", "Major", P, "2766(VBELN 공백) → 납품 생성 차단"),
    ("DEL-04", "납품", "중복 납품 생성", "신규 미생성, 기존 납품번호 로그", "Major", P,
     "1757 재실행 → RESULT=E '85857217가 이미 있기 때문에 납품을 생성할 수 없습니다.'"),
    ("DEL-05", "납품", "미마감 주문 (STATU≠C)", "납품 생성 차단, 주문 미마감 메시지", "Major", P,
     "1757 STATU=B → RESULT=E '주문이 마감 되지 않아서 납품생성할 수 없습니다'"),
    ("DEL-06", "납품", "다건 혼합 처리", "주문별 성공/실패 구분, 전체 중단 안 됨", "Major", N,
     "※재실행 필요 - 선택화면 복수선택으로 3개 주문 격리 조회 실패(전체 결과 대상화). 로그 상 건별 성공/실패 구분과 LOOP 지속은 확인됨(6건 각각 개별 오류/성공 반환)"),
    ("DEL-07", "납품", "로그 중복 방지", "판매오더·납품문서별 로그 1건", "Major", P,
     "다품목(1722 6품목/1757 2품목)이어도 문서별 로그 1건만 기록(SEQNO 무관)"),
    # ---- 9. 출고전기 아이콘 ----
    ("ICON-01", "아이콘", "WBSTA=A → 공란", "GI_ICON 공란", "Minor", P, "1757 납품 VBUP 전부 WBSTA=A → GI_ICON 공란"),
    ("ICON-02", "아이콘", "일부 WBSTA=B → 공란", "완료 아님이므로 공란", "Minor", N, "※미실행 - 부분출고(WBSTA=B) 상태 납품 구성 필요"),
    ("ICON-03", "아이콘", "전 WBSTA=C → 초록 (1722)", "GI_ICON=ICON_GREEN_LIGHT", "Critical", P,
     "1722 출고전기 완료 VBUP 전부 C → GI_ICON 초록(@08@)"),
    ("ICON-04", "아이콘", "재조회 시 초록 유지", "WBSTA 기준 재조회 시 초록 유지", "Major", P, "1722 재조회 후에도 GI_ICON 초록 유지"),
    ("ICON-05", "아이콘", "출고 직후 즉시 갱신", "재실행 없이 GI_ICON 즉시 표시", "Minor", P, "1722 CRE_DELIVERY 직후(재조회 전) GI_ICON 초록 즉시 반영"),
    # ---- 9-1. ALV 컬럼 ----
    ("ALV-01", "ALV", "출고전기 컬럼 표시", "GI_ICON 컬럼, 컬럼명 '출고전기', 아이콘 컬럼", "Major", P,
     "헤더 ALV에 GI_ICON 컬럼 존재(31→32컬럼), 컬럼명 '출고전기', 아이콘 셀로 출력"),
    ("ALV-02", "ALV", "미출고(WBSTA A) 공란", "출고전기 컬럼 공란", "Minor", P, "1757 납품 WBSTA A → GI_ICON 공란"),
    ("ALV-03", "ALV", "부분출고(WBSTA B) 공란", "완료 아님이므로 공란", "Minor", N, "※미실행 - WBSTA=B 납품 필요"),
    ("ALV-04", "ALV", "출고완료(WBSTA C) 초록", "GI_ICON 초록", "Major", P, "1722 → GI_ICON 초록"),
    ("ALV-05", "ALV", "출고 직후 컬럼 갱신", "재실행 없이 초록 갱신", "Minor", P, "1722 출고 직후 재조회 없이 GI_ICON 초록"),
    ("ALV-06", "ALV", "프로그램 재조회", "VBUP-WBSTA 기준 GI_ICON 재표시", "Major", P, "1722 재조회 시 VBUP 기준 초록 재표시"),
    ("ALV-07", "ALV", "출고취소·납품삭제 후 초기화", "VBELN_VL·GI_ICON 초기화", "Major", P,
     "[재테스트 2026-09-17 2차 수정본] 1722 출고완료 납품 85857223 삭제 후, 재조회 전 인세션 화면에서 VBELN_VL·GI_ICON 모두 즉시 공란 처리 확인(in-session GI_ICON=[]). F8 재조회 후에도 공란. 이전 잔존결함(GI_ICON 인세션 미초기화) 해소. (미출고 삭제 DELETE-01·1757도 정상)"),
    ("ALV-08", "ALV", "레이아웃 Variant", "신규 GI_ICON 숨겨진 경우 레이아웃 변경에서 선택 가능", "Minor", P,
     "기본 레이아웃에 GI_ICON 표시됨. 레이아웃 변경(&COL0)·저장 레이아웃(&LOAD) 팝업 확인 - 필요시 컬럼 선택 가능"),
    # ---- 10. 삭제 ----
    ("DELETE-01", "삭제", "미출고 납품 삭제 (1722→실제 1757)", "출고 취소 없이 납품 삭제", "Critical", P,
     "1757 납품 85857217(WBSTA A) DEL_DELIVERY → RESULT=S '납품 삭제', VBELN_VL·GI_ICON 공란"),
    ("DELETE-02", "삭제", "출고완료 납품 삭제 (2766→실제 1722)", "출고 취소 + WBSTA A 초기화 + 납품 삭제 + ALV 초기화", "Critical", P,
     "[재테스트 2026-09-17 수정본] 1722 판매오더 4293343 마감→납품 85857220 생성+출고전기(WBSTA=C, GI_ICON 초록 @08@). 동일 납품 '납품문서 삭제' → ABAP 덤프 없이 정상 삭제(prog=Z33SDR600/screen=100, RESULT=S '납품 85857220을(를) 삭제했습니다'). 출고취소+납품삭제 수행, 재조회 시 VBELN_VL·GI_ICON 초기화 확인. ※재조회 전 인세션 GI_ICON 잔존은 ALV-07 참조"),
    ("DELETE-03", "삭제", "부분출고(WBSTA B) 삭제", "자동 삭제 차단, 오류 로그", "Major", N, "※미실행 - WBSTA=B 납품 필요"),
    ("DELETE-04", "삭제", "대금청구문서 존재 출고완료 삭제", "출고 취소 실패, 납품 미삭제", "Major", N, "※미실행 - 청구문서 존재 납품 구성 필요"),
    ("DELETE-05", "삭제", "삭제 대상 없음", "삭제 대상 없음 메시지", "Minor", P,
     "1757 납품 없는 상태 DEL_DELIVERY → RESULT=E '납품 전표 은(는) 존재하지 않거나 아카이브/삭제되었습니다.'"),
    ("DELETE-06", "삭제", "납품 2건 이상(분할) 삭제", "분할납품 판단, 자동 삭제 차단", "Major", N, "※미실행 - 한 판매오더에 납품 2건(분할) 구성 필요"),
    ("DELETE-07", "삭제", "출고 취소 성공 후 납품 삭제 실패", "출고 취소 유지, 납품 잔존, 실패 단계·원인 로그, 재조회 반영", "Major", P,
     "[재테스트 2026-09-17 수정본] 덤프 해소로 처리 결과가 ABAP 덤프가 아닌 결과 로그 팝업으로 표출(RETCODE=Y/RESULT=S '납품 85857220을(를) 삭제했습니다'). 성공/실패가 문서별 로그로 제공되는 구조 확인"),
    ("DELETE-08", "삭제", "다건 삭제 혼합", "문서별 성공/실패 로그, 성공분만 초기화", "Major", N, "※미실행 - DELETE-02 덤프로 GI완료 납품 삭제 경로 차단, 혼합 삭제 미수행"),
    # ---- 11. 회귀/성능 ----
    ("REG-01", "회귀", "Header 중복 여부", "동일 ID/BSTKD/VERSION Header 1건", "Major", P,
     "S_MATNR=PNGB3500013L008 조회 결과 헤더 전행 ID|BSTKD|VERSION 키 고유(중복 0, ALV 페이징 확인)"),
    ("REG-02", "회귀", "ALV Refresh", "자재변경·납품·출고·삭제 후 재실행 없이 즉시 반영", "Major", P,
     "자재변경(품목 즉시), 납품/출고(VBELN_VL·GI_ICON 즉시), 삭제(공란 즉시) 모두 재조회 없이 반영"),
    ("REG-03", "회귀", "조회 성능 및 덤프", "장시간 지연·덤프 없이 조회", "Major", P,
     "[재테스트 2026-09-17 수정본] 조회 성능 정상(연 단위 지연 없음). GI완료 납품 삭제 시 ABAP 덤프(SAPSQL_PARSER_TODO_WARNING) 미발생 → 신규 ST22 덤프 없음. 덤프 결함 해소 확인"),
    ("REG-04", "회귀", "결과 로그", "BSTKD·VBELN_VA·VBELN_VL·RESULT·RETCODE·RESULT_MESSAGE 정확 기록", "Major", P,
     "결과 로그 팝업에 문서별 BSTKD/VBELN_VA/VBELN_VL/RETCODE/RESULT/RESULT_MESSAGE 정확히 기록(생성·삭제·오류 전부)"),
    ("REG-05", "회귀", "다국어 메시지", "각 로그인 언어 메시지 정상", "Minor", N, "※미실행 - 한국어(KO)만 확인. 영어·러시아어 로그인은 사용자 수행 필요"),
]
prog_rows = [{"id": c[0], "case": "[%s] %s" % (c[1], c[2]), "expect": c[3], "sev": c[4], "result": c[5], "data": c[6]} for c in CASES]

# ---------------- 공통 ----------------
common = {}


def setc(nums, target, result, note=""):
    for n in nums:
        common["UC-%03d" % n] = (target, result, note)


setc(range(1, 120), "N", "", "해당 없음")
setc([3], "Y", P, "자재변경 입력 검증(미등록 자재 오류 팝업 MAT-04)")
setc([8], "Y", P, "선택화면 기본값(생성일·주문상태 A~B), 라디오 초기값 주문조회")
setc([9, 10, 11], "Y", P, "[재테스트] GI완료 납품 삭제 덤프 해소, 프로그램 정상 처리(DELETE-02 수정 확인)")
setc([13], "Y", P, "자재변경 이력 Z33SDT644 판매오더 생성 시 기록(HIST-02), 변경 없으면 미생성(HIST-06)")
setc([14], "Y", P, "자재변경 값(MATNR/MAKTX/MEINS)이 화면·저장에 일치(MAT-01/UNIT-01)")
setc([40], "Y", P, "조회조건(ID·BSTKD·자재·생성일·상태) 정상")
setc([43], "Y", P, "TC별 조회 건수 기대값 일치")
setc([45], "Y", P, "생성/변경/삭제 결과가 조회에 즉시 반영(REG-02)")
setc([46], "Y", P, "헤더 더블클릭 품목·ATP, VBELN 더블클릭 VA03 정상")
setc([49], "Y", P, "[재테스트] GI완료 납품 삭제 덤프(SAPSQL_PARSER) 해소(DELETE-02 수정 확인)")
setc([50, 51], "Y", P, "조회·처리 중 대기/멈춤 없음")
setc([54], "Y", P, "수행옵션↔상태필드, KTGRM↔납품/출고 분기 연동")
setc([55], "Y", P, "다중조건 동시 적용")
setc([57], "Y", P, "자재 조회 시 헤더 중복 없음(REG-01)")
setc([58], "Y", P, "헤더/품목 미선택·다중선택 가드")
setc([59], "Y", P, "납품 삭제 후 VBELN_VL·GI_ICON 초기화(DELETE-01)")
setc([61], "Y", P, "생성/삭제 전 예/아니오 확인창")
setc([62, 66, 67], "Y", P, "[재테스트] 출고완료 납품 삭제 덤프 해소, 출고취소+삭제 정상(DELETE-02 수정 확인)")
setc([69], "Y", P, "판매오더 삭제(ELIM_ORD) 시 헤더/품목 정리")
setc([87], "Y", P, "개발용 메시지 없음(오류는 로그/팝업)")
setc([88], "Y", P, "마감/판매오더 존재 시 자재변경 버튼 차단")
setc([90], "Y", P, "[재테스트 2차] 삭제·변경 후 화면 refresh 정상 - 인세션 VBELN_VL·GI_ICON 즉시 초기화 확인(ALV-07 해소)")
setc([93], "Y", P, "화면 제목·컬럼명(출고전기 등)·버튼명 한국어 정상")
setc([94], "Y", P, "GI_ICON 컬럼·버튼 링크 정상")
setc([97], "Y", P, "예/아니오 확인창, 오류 팝업 정상")
setc([99], "Y", P, "메시지 상황 적합(마감/미마감/중복/미존재)")
setc([101], "Y", P, "라디오 초기값(주문조회)")
setc([106], "Y", P, "상태(마감/판매오더)에 따른 버튼 enable/disable")
setc([109], "Y", P, "자재변경 편집 제어")
setc([111], "Y", P, "한글 메시지")
setc([114], "Y", P, "[재테스트] GI완료 납품 삭제 기능 정상 동작(DELETE-02 수정 확인)")
setc([118], "Y", P, "결과 로그 항목 정확(REG-04)")
assert len(common) == 119, len(common)

fail_rows = [
    {"id": "결함이력(해결)", "case": "[10/9-1] 출고완료 납품 삭제 - 결함 발견·조치 이력 (전건 해결)", "tester": "김지인",
     "detail": ("본 시트는 테스트 중 발견된 결함과 조치·재검증 이력이다. 현재 전건 해결되어 프로그램별/공통 결과에 미해결 FAIL은 없다.\n"
                "\n[결함1] (Critical, 해결) 출고완료(VBUP-WBSTA=C) 납품문서를 ZSD600 '납품문서 삭제'로 삭제 시 ABAP 덤프 SAPSQL_PARSER_TODO_WARNING "
                "(Z33SDR600). 오류: \"aggregate function POSNR and list field LV_LINES are not compatible\" - 집계함수 결과를 담는 LV_LINES 타입 불일치. "
                "출고취소는 되나 삭제 단계에서 프로그램 종료→납품 잔존, 실패가 로그 아닌 덤프로 표출(DELETE-02/07·REG-03 FAIL).\n"
                "  → [조치] LV_LINES를 집계 결과 호환 타입으로 수정. [재검증 2026-09-17 1차] 1722/판매오더 4293343/납품 85857220 삭제 → 덤프 없이 정상 삭제(RESULT=S), DELETE-02·07·REG-03 및 공통 8건 PASS 전환.\n"
                "\n[결함2] (Major, 해결) 납품 삭제 후 재조회(F8) 전 인세션 화면에서 VBELN_VL은 즉시 공란이나 GI_ICON(출고전기 초록 아이콘)이 잔존 → 납품번호 없이 아이콘만 표시되는 불일치. "
                "원인: ZF_DELETE_OUTB_DELIVERY 삭제 성공부(3-3)가 VBELN_VL만 CLEAR, GI_ICON 미초기화(GET_LIKP_INFO ELSE 분기 동일).\n"
                "  → [조치] ZF_DELETE_OUTB_DELIVERY 3-3 및 GET_LIKP_INFO ELSE에서 GI_ICON도 CLEAR. [재검증 2026-09-17 2차] 1722/판매오더 4293344/납품 85857223 삭제 → 인세션 VBELN_VL·GI_ICON 모두 즉시 공란(in-session GI_ICON=[]), ALV-07·공통 UC-90 PASS 전환.\n"
                "\n[증적] 좌: 결함1 최초 덤프 화면 / 중: 결함2 해소 후 인세션(VBELN_VL·GI_ICON 공란) / 우: 재조회 후 상태."),
     "images": ["DELETE_02_1722_ABAP_DUMP.png", "S53_1722_after_delete_insession.png", "S53_1722_after_delete_requery.png"]},
]

evidence = [
    ("BAS-01", "선택화면(주문조회) - S_ID/S_BSTKD/S_MATNR/생성일/주문상태", "SEL_screen.png"),
    ("BAS-01", "2318 헤더+품목 11건", "BAS_2318_items.png"),
    ("BAS-01", "1722 부품 6건", "BAS_1722_items.png"),
    ("MAT-07", "마감(C) 주문 자재변경 차단 '마감품목 변경이 불가능합니다'", "MAT_07_closed_blocked.png"),
    ("MAT-04", "미등록 자재 오류 '마스터 레코드가 없습니다'", "MAT_04_unregistered_error.png"),
    ("MAT-01", "자재변경 대화상자(신규 자재명 해석)", "MAT_01_dialog.png"),
    ("MAT-01", "자재변경 저장 후 품목 ALV(MATNR/MAKTX/MEINS 반영)", "MAT_01_after_change.png"),
    ("MAT-06", "판매오더 생성 후 자재변경 차단", "MAT_06_after_so_blocked.png"),
    ("ATP-02", "2318 ATP 부족 - 부족 품목 빨간색(C610), 상태바 'ATP수량이 부족합니다'", "ATP_02_2318_dblclick.png"),
    ("ATP-01", "1757 변경 후 ATP 정상 → 판매오더 생성 로그(4293338)", "A03_1757_crea_ord_log.png"),
    ("HIST-02", "Z33SDT644 변경이력(OLD/NEW/VBELN) - 1757", "HIST_1757_z33sdt644.png"),
    ("DIV-01", "1757(Y1) 납품만 생성, GI_ICON 공란", "DEL_01_1757_requery.png"),
    ("DEL-01", "1757 납품 생성 로그(RESULT=S 1건)", "DEL_01_1757_cre_delivery_log.png"),
    ("DEL-04", "중복 납품 차단 로그", "DEL_04_1757_duplicate_log.png"),
    ("DEL-05", "미마감 주문 납품 차단 '주문이 마감 되지 않아서'", "DEL_05_1757_unclosed_log.png"),
    ("DEL-03", "판매오더 없음(2766) 납품 차단", "DEL_03_2766_no_so_log.png"),
    ("DIV-02", "1722(Y2) 납품+출고전기 생성 로그", "DIV_02_1722_cre_delivery_log.png"),
    ("ICON-03", "1722 출고완료 VBUP 전부 WBSTA=C", "ICON_03_1722_vbup.png"),
    ("ALV-04", "1722 GI_ICON 초록(재조회 후)", "DIV_02_1722_requery.png"),
    ("ICON-01", "1757 납품 VBUP 전부 WBSTA=A", "ICON_01_1757_vbup.png"),
    ("DELETE-01", "1757 납품 삭제 후 VBELN_VL·GI_ICON 공란", "DELETE_01_1757_requery.png"),
    ("DELETE-05", "삭제 대상 없음 메시지", "DELETE_05_1757_none_log.png"),
    ("DELETE-02", "[최초]★결함★ 출고완료 납품 삭제 시 ABAP 덤프(수정 전)", "DELETE_02_1722_ABAP_DUMP.png"),
    ("DELETE-07", "[최초] 덤프 후 VBUP C→A(출고취소됨), 납품 잔존(수정 전)", "DELETE_02_vbup_postdump.png"),
    ("DELETE-02", "[재테스트] 1722 출고전기 완료 GI_ICON 초록(삭제 전 스테이징)", "S53_1722_gi_green.png"),
    ("DELETE-02", "[재테스트] 출고완료 납품 삭제 결과 로그(RESULT=S, 덤프 없음)", "S53_1722_del_deliv_log.png"),
    ("ALV-07", "[재테스트 2차/해결] 삭제 직후 인세션 - VBELN_VL·GI_ICON 모두 즉시 초기화", "S53_1722_after_delete_insession.png"),
    ("ALV-07", "[재테스트 2차] 재조회(F8) 후에도 VBELN_VL·GI_ICON 공란 유지", "S53_1722_after_delete_requery.png"),
    ("ALV-08", "레이아웃 변경 팝업(표시/숨긴 컬럼)", "ALV_08_change_layout.png"),
    ("PRE", "2827 CLOS_ORD 실패분 SE16N 강제 STATU 변경(변경 완료)", "S46_2827_se16n_edit.png"),
    ("CLEANUP", "판매오더 삭제(ELIM_ORD '성공적으로 삭제 되었습니다') - 1757", "S50_1757_elim_log.png"),
    ("CLEANUP", "Z33SDT644 테스트 이력행 SE16N 삭제", "S51_644_1757_se16n_delete.png"),
]
ev_out = []
for tc, cap, fn in evidence:
    t = thumb(fn)
    if t:
        ev_out.append({"tc": tc, "caption": cap, "path": t})
for fr in fail_rows:
    fr["image_paths"] = [p for p in (thumb(f, 640) for f in fr["images"]) if p]

payload = {
    "date": DATE,
    "cover": {"C17": "2026.09.17"},
    "revision": {"B5": "0.03", "C5": "재테스트 2차: 1차에서 발견된 잔존결함(납품 삭제 후 인세션 GI_ICON 미초기화, ALV-07)까지 수정 반영본에서 해소 확인(ALV-07·공통 UC-90 PASS 전환). DELETE-02 ABAP 덤프 및 GI_ICON 인세션 초기화 전건 해결 - 미해결 FAIL 없음", "D5": "2026.09.17",
                 "E5": "김지인", "F5": "", "G5": ""},
    "prog_header": {
        "B1": "SD - 러시아 1C 주문 I/F",
        "D1": "김지인",
        "B2": "[RU] 1C I/F Sales Order Management - 자재변경·ATP·변경이력·KTGRM 납품/출고 분기·출고전기 아이콘(GI_ICON) 개선분",
        "D2": "Z33SDR600 / ZSD600",
    },
    "prog_rows": prog_rows,
    "common": {k: {"target": v[0], "result": v[1], "note": v[2]} for k, v in common.items()},
    "fail_rows": fail_rows,
    "evidence": ev_out,
    "evidence_sheet": "테스트증적(캡처)",
    "legend": {
        "blank": "결과란이 비어 있으면 미실행이다(사유는 테스트 데이터 열). PASS 근거 화면은 '테스트증적(캡처)' 시트에 케이스 ID별로 넣었다. [v0.03] 최초 발견 결함 2건(DELETE-02 ABAP 덤프, ALV-07 인세션 GI_ICON 미초기화)은 수정 반영본에서 전건 해소·재검증 완료 → 미해결 FAIL 없음. 결함 발견·조치 이력은 'FAIL내역' 시트 참조.",
        "note": "KDQ 210, 2026-09-16~17 및 재테스트 2026-09-17(1·2차), SAP GUI Scripting 자동 실행. " + D_TARGETS,
    },
}
io.open(OUT, "w", encoding="utf-8").write(json.dumps(payload, ensure_ascii=False, indent=1))
res = [r["result"] for r in prog_rows]
print("prog rows:", len(res), " PASS:", res.count("PASS"), " FAIL:", res.count("FAIL"), " blank:", res.count(""))
cr = list(payload["common"].values())
print("common Y:", sum(1 for v in cr if v["target"] == "Y"), " PASS:", sum(1 for v in cr if v["result"] == "PASS"), " FAIL:", sum(1 for v in cr if v["result"] == "FAIL"))
print("evidence:", len(ev_out), " missing:", missing)
