# make_xlsx.py : build the Excel deliverable for the ZCRM052 vs ZCRM052_S comparison test
import io, os, glob, collections
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

RAW = r"C:\sap-test\results\raw"
OUT = r"C:\sap-test\results\ZCRM052_vs_ZCRM052S_테스트결과_20260910.xlsx"

FONT = "Arial"
H_FILL = PatternFill("solid", fgColor="1F3864")
H_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=10)
SUB_FILL = PatternFill("solid", fgColor="D9E1F2")
SUB_FONT = Font(name=FONT, bold=True, size=10)
BODY = Font(name=FONT, size=10)
BOLD = Font(name=FONT, bold=True, size=10)
TITLE = Font(name=FONT, bold=True, size=14)
NOTE = Font(name=FONT, size=9, italic=True, color="666666")
OK_FILL = PatternFill("solid", fgColor="E2EFDA")
NG_FILL = PatternFill("solid", fgColor="FCE4D6")
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def load_tsv(path):
    with io.open(path, encoding="utf-16") as fh:
        lines = fh.read().split("\n")
    idx = [n for n, l in enumerate(lines) if l.startswith("#DATA")]
    meta = {}
    for l in lines:
        p = l.rstrip("\r").split("\t")
        if p[0].startswith("#") and p[0] != "#DATA" and len(p) > 1:
            meta.setdefault(p[0].lstrip("#"), p[1])
    if not idx:
        return meta, [], []
    i = idx[0]
    hdr = lines[i + 1].rstrip("\r").split("\t")
    rows = [l.rstrip("\r").split("\t") for l in lines[i + 2:] if l.strip()]
    return meta, hdr, rows


def cols_of(path):
    """technical column list in display order from an S01 dump"""
    out = []
    with io.open(path, encoding="utf-16") as fh:
        for l in fh.read().split("\n"):
            p = l.rstrip("\r").split("\t")
            if p[0] == "#COL" and len(p) > 2:
                out.append((p[2], p[3] if len(p) > 3 else ""))
    return out


def style_header(ws, row, ncol, start=1):
    for c in range(start, start + ncol):
        cell = ws.cell(row=row, column=c)
        cell.fill = H_FILL
        cell.font = H_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BOX


def put_row(ws, row, values, font=BODY, fill=None, start=1):
    for j, v in enumerate(values):
        cell = ws.cell(row=row, column=start + j, value=v)
        cell.font = font
        cell.border = BOX
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        if fill:
            cell.fill = fill
    return row + 1


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


wb = Workbook()

# ---------------------------------------------------------------- 1. 개요
ws = wb.active
ws.title = "1.개요"
ws["A1"] = "ZCRM052 vs ZCRM052S 비교 테스트 결과"
ws["A1"].font = TITLE
ws["A2"] = "회수품 적합성 관리 - 구독 전용 프로그램 분리 회귀 테스트"
ws["A2"].font = NOTE

r = 4
r = put_row(ws, r, ["항목", "내용"], H_FONT, H_FILL)
info = [
    ("테스트 일자", "2026-09-10"),
    ("대상 시스템", "KDQ (QA) / 클라이언트 210"),
    ("사용자", "KKR003-1 / 언어 KO"),
    ("기존 프로그램", "T-code ZCRM052 → 프로그램 ZCRMR052"),
    ("신규 프로그램", "T-code ZCRM052S → 프로그램 ZCRMR052S"),
    ("조회 기간", "2026.07.01 ~ 2026.07.15 (15일)"),
    ("구독조건 상태", "조회 SELECT 2곳 주석 처리됨 / 저장 경로 구독 검증 제거됨"),
    ("테스트 방식", "SAP GUI Scripting 자동화. 두 프로그램에 동일 조건 적용 후 결과 대조"),
    ("테스트 식별자", "ZTEST_20260909 (저장 4건 모두 이 값)"),
]
for k, v in info:
    r = put_row(ws, r, [k, v])

r += 1
ws.cell(row=r, column=1, value="※ 'ZCRM052_S' 는 이 문서의 표기다. 실제 임시 T-code 값은 문서에 적지 않는다.").font = NOTE
r += 2
ws.cell(row=r, column=1, value="시트 구성").font = BOLD
r += 1
r = put_row(ws, r, ["시트", "내용"], H_FONT, H_FILL)
for k, v in [
    ("2.결과요약", "케이스별 판정 일람 및 집계"),
    ("3.조회건수비교", "TC-01~05 건수 대조 및 검산"),
    ("4.ALV컬럼비교", "표시 컬럼 전수 대조"),
    ("5.검증메시지비교", "TC-09 저장 검증 6종"),
    ("6.편집제한비교", "TC-12 셀 편집 가능 여부"),
    ("7.상세조회비교", "TC-15 팝업 항목·값"),
    ("8.저장결과비교", "TC-07·TC-08 저장 결과 및 관련 테이블"),
    ("9.사용데이터", "사용한 실제 데이터와 정리 기준"),
    ("10.미해결·확인요청", "개발자 확인 사항"),
    ("11.화면캡처목록", "증거 이미지 파일 목록"),
    ("12.주요화면", "핵심 캡처 삽입"),
]:
    r = put_row(ws, r, [k, v])
widths(ws, {"A": 22, "B": 78})

# ---------------------------------------------------------------- 2. 결과요약
ws = wb.create_sheet("2.결과요약")
ws["A1"] = "테스트 결과 요약"
ws["A1"].font = TITLE
r = 3
hdr = ["TC ID", "테스트 항목", "기존 ZCRM052 결과", "신규 ZCRM052S 결과", "일치 여부", "비고"]
r = put_row(ws, r, hdr, H_FONT, H_FILL)
summary = [
    ("TC-01", "센터조치 접수 조회", "5,312건", "5,312건", "일치", "어제와 동일"),
    ("TC-02", "대리점조치 접수 조회", "1,041건", "1,041건", "일치", "어제와 동일"),
    ("TC-03", "회수품 미접수 조회", "395건", "395건", "일치", "395행 전 컬럼 전수 비교 일치"),
    ("TC-04", "유상/무상 조회", "유상 3,460 / 무상 1,852", "유상 3,460 / 무상 1,852", "일치", "합계가 전체와 일치"),
    ("TC-05", "멀티킷 조회", "O 24 / X 5,288", "O 24 / X 5,288", "일치", "화면 O/X → DB X/공백 변환 동일"),
    ("TC-06", "ALV 화면 구성", "42개 컬럼 표시", "44개 컬럼 표시", "조건부 일치",
     "레이아웃 유무 차이. NCF_RSN_V 미표시는 어제부터 미해결"),
    ("TC-07", "적합수량 저장", "접수번호 2026091000001", "접수번호 2026091000002", "일치",
     "어제 실행 불가 → 오늘 정상 저장. 저장 항목 채움 규칙 동일"),
    ("TC-08", "부적합수량 저장", "접수번호 2026091000003", "접수번호 2026091000004", "일치",
     "NCF=N, 사유 Z02 저장 동일"),
    ("TC-09", "저장 검증 메시지", "6종 정상 차단", "6종 정상 차단", "조건부 일치",
     "5종 문구까지 동일. TC-09-06 만 마침표·검증시점 차이"),
    ("TC-10", "책임사업장 F4", "-", "-", "미실행", "ALV 셀 F4 팝업이 열리지 않음"),
    ("TC-11", "부적합사유 Dropdown", "셀 타입 ValueList", "셀 타입 ValueList", "부분 일치",
     "저장 시 Z02 정상 반영 확인. 목록 내용은 미확인"),
    ("TC-12", "품질검사 완료 행 잠금", "잠금 규칙 2종", "동일", "일치", "짝지은 49건 불일치 0건"),
    ("TC-13", "해지점 자동 부적합", "-", "-", "데이터 없음", "4~7월 4개 구간 대상 0건"),
    ("TC-14", "단납 보증기간 계산", "보증기간 60 등 일치", "동일", "부분 일치", "일반 보증 경로만 확인"),
    ("TC-15", "상세조회", "10개 항목", "동일", "일치", "컬럼과 값 모두 일치"),
    ("TC-16", "Header 변경 감지", "-", "-", "미실행", ""),
    ("TC-17", "변경 없이 종료", "팝업 없음, 정상 종료", "동일", "일치", ""),
]
first = r
for row in summary:
    fill = None
    if row[4] == "일치":
        fill = OK_FILL
    elif row[4] in ("실행 불가", "비교 불가", "데이터 없음", "미실행"):
        fill = NG_FILL
    else:
        fill = WARN_FILL
    r = put_row(ws, r, list(row), BODY, fill)
last = r - 1

r += 1
ws.cell(row=r, column=1, value="집계").font = BOLD
r += 1
r = put_row(ws, r, ["판정", "건수"], H_FONT, H_FILL)
for label in ["일치", "조건부 일치", "부분 일치", "비교 불가", "실행 불가", "데이터 없음", "미실행"]:
    ws.cell(row=r, column=1, value=label).font = BODY
    ws.cell(row=r, column=1).border = BOX
    ws.cell(row=r, column=2, value='=COUNTIF($E${}:$E${},A{})'.format(first, last, r)).font = BODY
    ws.cell(row=r, column=2).border = BOX
    r += 1
ws.cell(row=r, column=1, value="합계").font = BOLD
ws.cell(row=r, column=1).border = BOX
ws.cell(row=r, column=2, value="=SUM(B{}:B{})".format(r - 7, r - 1)).font = BOLD
ws.cell(row=r, column=2).border = BOX
widths(ws, {"A": 10, "B": 26, "C": 26, "D": 26, "E": 14, "F": 42})
ws.freeze_panes = "A4"

# ---------------------------------------------------------------- 3. 조회건수비교
ws = wb.create_sheet("3.조회건수비교")
ws["A1"] = "조회 건수 비교 (2026.07.01 ~ 2026.07.15)"
ws["A1"].font = TITLE
r = 3
r = put_row(ws, r, ["케이스", "처리조직", "접수여부", "추가조건", "날짜필드",
                    "기존 건수", "신규 건수", "일치 여부"], H_FONT, H_FILL)
cases = [
    ("TC-01", "센터조치", "회수품 접수", "-", "접수일자", 5312, 5312),
    ("TC-02", "서비스점조치", "회수품 접수", "-", "접수일자", 1041, 1041),
    ("TC-03", "센터조치", "회수품 미접수", "-", "회수품요청일자", 395, 395),
    ("TC-04Y", "센터조치", "회수품 접수", "유/무상 = Y", "접수일자", 3460, 3460),
    ("TC-04N", "센터조치", "회수품 접수", "유/무상 = N", "접수일자", 1852, 1852),
    ("TC-05O", "센터조치", "회수품 접수", "멀티킷 = O", "접수일자", 24, 24),
    ("TC-05X", "센터조치", "회수품 접수", "멀티킷 = X", "접수일자", 5288, 5288),
]
start = r
for c in cases:
    ws.cell(row=r, column=1, value=c[0]).font = BODY
    ws.cell(row=r, column=2, value=c[1]).font = BODY
    ws.cell(row=r, column=3, value=c[2]).font = BODY
    ws.cell(row=r, column=4, value=c[3]).font = BODY
    ws.cell(row=r, column=5, value=c[4]).font = BODY
    ws.cell(row=r, column=6, value=c[5]).font = BODY
    ws.cell(row=r, column=7, value=c[6]).font = BODY
    ws.cell(row=r, column=8, value='=IF(F{r}=G{r},"일치","불일치")'.format(r=r)).font = BOLD
    for col in range(1, 9):
        ws.cell(row=r, column=col).border = BOX
    ws.cell(row=r, column=6).number_format = "#,##0"
    ws.cell(row=r, column=7).number_format = "#,##0"
    ws.cell(row=r, column=8).alignment = Alignment(horizontal="center")
    r += 1
end = r - 1

r += 1
ws.cell(row=r, column=1, value="검산 - 분기 합계가 전체와 일치하는지").font = BOLD
r += 1
r = put_row(ws, r, ["검산 항목", "계산식", "기존", "신규", "판정"], H_FONT, H_FILL)
checks = [
    ("유상 + 무상", "TC-04Y + TC-04N", start + 3, start + 4, start),
    ("멀티킷 O + X", "TC-05O + TC-05X", start + 5, start + 6, start),
]
for label, expr, a, b, tot in checks:
    ws.cell(row=r, column=1, value=label).font = BODY
    ws.cell(row=r, column=2, value=expr).font = BODY
    ws.cell(row=r, column=3, value="=F{}+F{}".format(a, b)).font = BODY
    ws.cell(row=r, column=4, value="=G{}+G{}".format(a, b)).font = BODY
    ws.cell(row=r, column=5,
            value='=IF(AND(C{r}=$F${t},D{r}=$G${t}),"전체와 일치","불일치")'.format(r=r, t=tot)).font = BOLD
    for col in range(1, 6):
        ws.cell(row=r, column=col).border = BOX
    ws.cell(row=r, column=3).number_format = "#,##0"
    ws.cell(row=r, column=4).number_format = "#,##0"
    r += 1

r += 1
ws.cell(row=r, column=1,
        value="※ 조회 1건당 6~34초 소요. 건수에 비례한다. 로그: logs/S01_compare_query.txt").font = NOTE
widths(ws, {"A": 12, "B": 16, "C": 16, "D": 16, "E": 18, "F": 12, "G": 12, "H": 12})
ws.freeze_panes = "A4"

# ---------------------------------------------------------------- 4. ALV컬럼비교
ws = wb.create_sheet("4.ALV컬럼비교")
ws["A1"] = "ALV 표시 컬럼 비교 (TC-06)"
ws["A1"].font = TITLE
ws["A2"] = "ColumnOrder 는 화면에 표시 중인 컬럼만 반환한다. 숨김 컬럼은 포함되지 않는다."
ws["A2"].font = NOTE

old_cols = cols_of(os.path.join(RAW, "TC01_ZCRMR052.tsv"))
new_cols = cols_of(os.path.join(RAW, "TC01_ZCRMR052S.tsv"))
old_names = [c for c, _ in old_cols]
new_names = [c for c, _ in new_cols]

r = 4
r = put_row(ws, r, ["구분", "기존 ZCRMR052", "신규 ZCRMR052S"], H_FONT, H_FILL)
r = put_row(ws, r, ["표시 컬럼 수", len(old_cols), len(new_cols)])
r = put_row(ws, r, ["저장된 ALV 레이아웃", "6개 (/CRM01 이 기본 설정)", "0개"])
r += 1

ws.cell(row=r, column=1, value="차이 컬럼").font = BOLD
r += 1
r = put_row(ws, r, ["구분", "컬럼", "의미", "판정"], H_FONT, H_FILL)
only_old = [c for c in old_names if c not in new_names]
only_new = [c for c in new_names if c not in old_names]
mean = {"NCF_RSN_V": "부적합사유 코드", "SUBS_TYPE": "구독 구분",
        "STATUST": "상태", "MULTIKITYN": "멀티킷 사용여부"}
for c in only_old:
    r = put_row(ws, r, ["기존에만 표시", c, mean.get(c, ""), "개발자 확인 필요"], BODY, WARN_FILL)
for c in only_new:
    verdict = "비교 제외 항목" if c == "SUBS_TYPE" else "레이아웃 차이로 설명됨"
    r = put_row(ws, r, ["신규에만 표시", c, mean.get(c, ""), verdict], BODY,
                OK_FILL if c != "SUBS_TYPE" else None)
r += 1

ws.cell(row=r, column=1, value="전체 컬럼 대조 - 표시 순서대로").font = BOLD
r += 1
head = r
r = put_row(ws, r, ["순번", "기존 컬럼", "기존 컬럼명", "신규에 존재", "순번", "신규 컬럼",
                    "신규 컬럼명", "기존에 존재"], H_FONT, H_FILL)
n = max(len(old_cols), len(new_cols))
body_start = r
for i in range(n):
    if i < len(old_cols):
        ws.cell(row=r, column=1, value=i).font = BODY
        ws.cell(row=r, column=2, value=old_cols[i][0]).font = BODY
        ws.cell(row=r, column=3, value=old_cols[i][1].split("|")[0]).font = BODY
        ws.cell(row=r, column=4,
                value='=IF(ISNUMBER(MATCH(B{r},$F${s}:$F${e},0)),"O","X")'.format(
                    r=r, s=body_start, e=body_start + n - 1)).font = BODY
        ws.cell(row=r, column=4).alignment = Alignment(horizontal="center")
    if i < len(new_cols):
        ws.cell(row=r, column=5, value=i).font = BODY
        ws.cell(row=r, column=6, value=new_cols[i][0]).font = BODY
        ws.cell(row=r, column=7, value=new_cols[i][1].split("|")[0]).font = BODY
        ws.cell(row=r, column=8,
                value='=IF(ISNUMBER(MATCH(F{r},$B${s}:$B${e},0)),"O","X")'.format(
                    r=r, s=body_start, e=body_start + n - 1)).font = BODY
        ws.cell(row=r, column=8).alignment = Alignment(horizontal="center")
    for col in range(1, 9):
        ws.cell(row=r, column=col).border = BOX
    r += 1
ws.cell(row=head - 1, column=1).comment = Comment(
    "O = 상대 프로그램에도 표시되는 컬럼, X = 한쪽에만 표시되는 컬럼", "test")
widths(ws, {"A": 7, "B": 18, "C": 22, "D": 12, "E": 7, "F": 18, "G": 22, "H": 12})
ws.freeze_panes = "A{}".format(body_start)

# ---------------------------------------------------------------- 5. 검증메시지비교
ws = wb.create_sheet("5.검증메시지비교")
ws["A1"] = "저장 검증 메시지 비교 (TC-09)"
ws["A1"].font = TITLE
ws["A2"] = "저장 버튼 → '저장 하시겠습니까?' 팝업에서 예 → 프로그램 검증 실행"
ws["A2"].font = NOTE
r = 4
r = put_row(ws, r, ["케이스", "입력 조건", "기존 ZCRM052 메시지",
                    "신규 ZCRM052S 메시지", "일치 여부", "비고"], H_FONT, H_FILL)
msgs = [
    ("TC-09-01", "적합수량 6 (요청수량 1)",
     "적합/부적합 수량의 합은 요청수량을 초과 할 수 없습니다.",
     "적합/부적합 수량의 합은 요청수량을 초과 할 수 없습니다.", ""),
    ("TC-09-02", "부적합수량 1, 사유 공백",
     "부적합 사유가 누락 되었습니다.", "부적합 사유가 누락 되었습니다.", ""),
    ("TC-09-03", "부적합수량 0, 사유 Z02 입력",
     "부적합 수량이 없는데, 부접합 사유가 있습니다.",
     "부적합 수량이 없는데, 부접합 사유가 있습니다.", ""),
    ("TC-09-04", "적합수량 1, 책임사업장 공백",
     "적합 수량이 존재 하면, 책임사업장은 필 수 입니다.",
     "적합 수량이 존재 하면, 책임사업장은 필 수 입니다.", ""),
    ("TC-09-05", "접수부서·접수자 공백",
     "접수부서명 또는 접수자가 누락 되었습니다.",
     "접수부서명 또는 접수자가 누락 되었습니다.", ""),
    ("TC-09-06", "회수품 입고일자 공백",
     "회수품 입고일자를 입력해주세요",
     "회수품 입고일자를 입력해주세요.",
     "신규는 문장 끝에 마침표가 있고, 검증 시점도 다르다"),
]
firstm = r
for cid, cond, om, nm, note in msgs:
    ws.cell(row=r, column=1, value=cid).font = BODY
    ws.cell(row=r, column=2, value=cond).font = BODY
    ws.cell(row=r, column=3, value=om).font = BODY
    ws.cell(row=r, column=4, value=nm).font = BODY
    ws.cell(row=r, column=5, value='=IF(EXACT(C{r},D{r}),"일치","불일치")'.format(r=r)).font = BOLD
    ws.cell(row=r, column=6, value=note).font = BODY
    fill = OK_FILL if om == nm else NG_FILL
    for c2 in range(1, 7):
        ws.cell(row=r, column=c2).border = BOX
        ws.cell(row=r, column=c2).fill = fill
        ws.cell(row=r, column=c2).alignment = Alignment(vertical="center", wrap_text=True)
    ws.cell(row=r, column=5).alignment = Alignment(horizontal="center")
    r += 1
lastm = r - 1
r += 1
ws.cell(row=r, column=1, value="일치 건수").font = BOLD
ws.cell(row=r, column=2, value='=COUNTIF(E{}:E{},"일치")'.format(firstm, lastm)).font = BOLD
r += 1
ws.cell(row=r, column=1, value="불일치 건수").font = BOLD
ws.cell(row=r, column=2, value='=COUNTIF(E{}:E{},"불일치")'.format(firstm, lastm)).font = BOLD
r += 2
ws.cell(row=r, column=1, value="TC-09-06 차이 상세 - 개발자 확인 필요").font = BOLD
r += 1
r = put_row(ws, r, ["항목", "기존 ZCRMR052", "신규 ZCRMR052S"], H_FONT, H_FILL)
r = put_row(ws, r, ["메시지", "회수품 입고일자를 입력해주세요",
                    "회수품 입고일자를 입력해주세요. (마침표 추가)"], BODY, NG_FILL)
r = put_row(ws, r, ["검증 시점", "저장 버튼 즉시 차단. 확인 팝업 없음",
                    "저장 확인 팝업 → 예 → 차단"], BODY, NG_FILL)
r += 1
ws.cell(row=r, column=1, value=(
    "2026-09-09 에는 신규가 '구독 대상이 아닌 데이터가 포함되어 있습니다.' 로 6종 전부 "
    "선차단되어 비교가 불가능했다. 저장 경로의 구독 검증이 제거되어 오늘 비교가 가능해졌다.")).font = BODY
widths(ws, {"A": 12, "B": 26, "C": 44, "D": 44, "E": 11, "F": 34})
ws.freeze_panes = "A5"
ws = wb.create_sheet("6.편집제한비교")
ws["A1"] = "셀 편집 제한 비교 (TC-12)"
ws["A1"].font = TITLE
ws["A2"] = "접수일자 2026.07.01 조회(양쪽 2,922건) 중 상위 401행을 읽고 같은 레코드끼리 짝지어 비교"
ws["A2"].font = NOTE
r = 4
r = put_row(ws, r, ["항목", "값"], H_FONT, H_FILL)
for k, v in [("양쪽 조회 건수", "2,922건 / 2,922건"),
             ("각 프로그램 수집 행", "401행"),
             ("짝지어 비교한 레코드", "49건"),
             ("편집 가능 여부 불일치", "0건")]:
    r = put_row(ws, r, [k, v], BODY, OK_FILL if "0건" in str(v) else None)
r += 1
ws.cell(row=r, column=1, value="관측된 잠금 패턴 (두 프로그램 동일)").font = BOLD
r += 1
r = put_row(ws, r, ["패턴", "적합수량", "부적합수량", "부적합사유", "책임사업장", "비고",
                    "회수품입고일자", "건수", "의미"], H_FONT, H_FILL)
r = put_row(ws, r, ["전체 편집 가능", "가능", "가능", "가능", "가능", "가능", "가능", 30,
                    "미접수·미완료 건"], BODY, OK_FILL)
r = put_row(ws, r, ["접수 완료 잠금", "잠김", "잠김", "잠김", "잠김", "잠김", "가능", 19,
                    "접수 완료 건"], BODY, OK_FILL)
r += 1
ws.cell(row=r, column=1, value=(
    "※ 미검증: ZCRM_QRST_PRT-ZCLOSE = 'X' 인 품질검사 완전 마감 건은 2026년 데이터에 없다. "
    "접수번호 2026 전체 범위 조회 결과 0건.")).font = NOTE
widths(ws, {"A": 18, "B": 12, "C": 12, "D": 12, "E": 12, "F": 10, "G": 16, "H": 8, "I": 22})

# ---------------------------------------------------------------- 7. 상세조회비교
ws = wb.create_sheet("7.상세조회비교")
ws["A1"] = "상세조회 비교 (TC-15) 와 종료 동작 (TC-17)"
ws["A1"].font = TITLE
r = 3
r = put_row(ws, r, ["조작", "기존 ZCRM052", "신규 ZCRM052S", "일치 여부"], H_FONT, H_FILL)
for a, b, c in [
    ("행 미선택 후 상세조회", "상태바 '하나의 라인을 선택 하세요 .' (E)", "동일"),
    ("1행 선택 후 상세조회", "보증기간 ALV 팝업 표시", "동일"),
    ("2행 선택 후 상세조회", "상태바 '하나의 라인을 선택 하세요 .' (E)", "동일"),
    ("변경 없이 뒤로가기 (F3)", "팝업 없음, 선택화면 복귀", "동일"),
]:
    r = put_row(ws, r, [a, b, c, "일치"], BODY, OK_FILL)
r += 1
ws.cell(row=r, column=1, value="상세조회 팝업 항목 - 컬럼과 값 모두 일치").font = BOLD
r += 1
r = put_row(ws, r, ["컬럼", "의미", "기존 값", "신규 값", "일치 여부"], H_FONT, H_FILL)
detail = [
    ("MATNR", "자재", "PEME521QP001"),
    ("MATNRT", "자재명", "전기매트_세트;EME521,Pad,Queen,23년형"),
    ("PRODH", "제품 계층", "110500002000010"),
    ("PRODHT", "제품 계층명", "EME"),
    ("PRDDATE", "제조일", "2023.11.02"),
    ("ISTDATE", "설치일", "2024.01.02"),
    ("PART_PERIOD", "보증기간", "60"),
    ("M_FROM", "제조 기준 시작일", "2021.09.01"),
    ("IST_FROM", "설치 기준 시작일", "2021.09.01"),
    ("IST_TO", "설치 기준 종료일", "3000.12.31"),
]
for col, mn, val in detail:
    ws.cell(row=r, column=1, value=col).font = BODY
    ws.cell(row=r, column=2, value=mn).font = BODY
    ws.cell(row=r, column=3, value=val).font = BODY
    ws.cell(row=r, column=4, value=val).font = BODY
    ws.cell(row=r, column=5, value='=IF(C{r}=D{r},"일치","불일치")'.format(r=r)).font = BOLD
    for c2 in range(1, 6):
        ws.cell(row=r, column=c2).border = BOX
        ws.cell(row=r, column=c2).fill = OK_FILL
    ws.cell(row=r, column=5).alignment = Alignment(horizontal="center")
    r += 1
r += 1
ws.cell(row=r, column=1, value=(
    "※ 보증기간과 기준일자가 값까지 일치하므로 TC-14 의 일반 보증 경로도 이 범위에서 동일함이 확인된다. "
    "단납(제조번호 기반) 경로는 별도 데이터가 필요하다.")).font = NOTE
widths(ws, {"A": 16, "B": 20, "C": 40, "D": 40, "E": 12})

# ---------------------------------------------------------------- 8. 사용데이터
ws = wb.create_sheet("8.저장결과비교")
ws["A1"] = "저장 결과 비교 (TC-07 적합수량 / TC-08 부적합수량)"
ws["A1"].font = TITLE
ws["A2"] = "두 프로그램 모두 '정상적으로 저장되었습니다.' 로 저장 완료"
ws["A2"].font = NOTE
r = 4
ws.cell(row=r, column=1, value="TC-07 적합수량 저장").font = BOLD
r += 1
r = put_row(ws, r, ["항목", "기존 ZCRMR052", "신규 ZCRMR052S", "일치 여부"], H_FONT, H_FILL)
tc07 = [
    ("상담번호", "2026070703501", "2026062402790", "-"),
    ("자재번호(부품)", "30036673A", "30030863A", "-"),
    ("꼬리표NO", "0010", "0093", "-"),
    ("채번된 접수번호", "2026091000001", "2026091000002", "-"),
    ("접수품목번호", "10", "10", "chk"),
    ("접수일자", "2026.09.10", "2026.09.10", "chk"),
    ("적합수량 CF_QTY", "1", "1", "chk"),
    ("부적합수량 NCF_QTY", "0", "0", "chk"),
    ("적합여부 NCF", "Y", "Y", "chk"),
    ("책임사업장 RSPC", "1100", "1100", "chk"),
    ("상태 STATUS", "X", "X", "chk"),
    ("적합판정일 CONFDAT", "20260910", "20260910", "chk"),
    ("금액 NETWR", "2,500", "61,500", "-"),
    ("통화 WAERS", "KRW", "KRW", "chk"),
    ("회수품 입고일자", "2026.07.15", "2026.07.15", "chk"),
]
for nm, a, b, mode in tc07:
    ws.cell(row=r, column=1, value=nm).font = BODY
    ws.cell(row=r, column=2, value=a).font = BODY
    ws.cell(row=r, column=3, value=b).font = BODY
    if mode == "chk":
        ws.cell(row=r, column=4, value='=IF(EXACT(B{r},C{r}),"일치","불일치")'.format(r=r)).font = BOLD
        ws.cell(row=r, column=4).fill = OK_FILL
    else:
        ws.cell(row=r, column=4, value="데이터 상이").font = BODY
    for c2 in range(1, 5):
        ws.cell(row=r, column=c2).border = BOX
    r += 1
ws.cell(row=r, column=1, value=(
    "금액은 두 건의 자재가 달라 단가가 다르다. 저장 항목의 채움 규칙은 동일하다.")).font = NOTE
r += 2
ws.cell(row=r, column=1, value="TC-08 부적합수량 저장").font = BOLD
r += 1
r = put_row(ws, r, ["항목", "기존 ZCRMR052", "신규 ZCRMR052S", "일치 여부"], H_FONT, H_FILL)
tc08 = [
    ("상담번호", "2026070600189", "2026070202547", "-"),
    ("자재번호(부품)", "20050715A", "30019272A", "-"),
    ("꼬리표NO", "0201", "0205", "-"),
    ("채번된 접수번호", "2026091000003", "2026091000004", "-"),
    ("적합수량 CF_QTY", "0", "0", "chk"),
    ("부적합수량 NCF_QTY", "1", "1", "chk"),
    ("적합여부 NCF", "N", "N", "chk"),
    ("부적합사유 NCF_RSN_V", "Z02", "Z02", "chk"),
    ("책임사업장 RSPC", "(공백)", "(공백)", "chk"),
    ("상태 STATUS", "X", "X", "chk"),
    ("적합판정일 CONFDAT", "(공백)", "(공백)", "chk"),
    ("금액 NETWR", "(공백)", "(공백)", "chk"),
    ("회수품 입고일자", "2026.07.15", "2026.07.15", "chk"),
]
for nm, a, b, mode in tc08:
    ws.cell(row=r, column=1, value=nm).font = BODY
    ws.cell(row=r, column=2, value=a).font = BODY
    ws.cell(row=r, column=3, value=b).font = BODY
    if mode == "chk":
        ws.cell(row=r, column=4, value='=IF(EXACT(B{r},C{r}),"일치","불일치")'.format(r=r)).font = BOLD
        ws.cell(row=r, column=4).fill = OK_FILL
    else:
        ws.cell(row=r, column=4, value="데이터 상이").font = BODY
    for c2 in range(1, 5):
        ws.cell(row=r, column=c2).border = BOX
    r += 1
r += 1
ws.cell(row=r, column=1, value="관련 테이블 검증").font = BOLD
r += 1
r = put_row(ws, r, ["테이블", "확인 내용", "결과"], H_FONT, H_FILL)
r = put_row(ws, r, ["ZCRM_RECPT", "접수일자 2026.09.10 조회", "정확히 4건"], BODY, OK_FILL)
r = put_row(ws, r, ["ZCRM_RECPT_WEAR", "접수번호 2026091000001~4",
                    "4건, 회수품 입고일자 2026.07.15"], BODY, OK_FILL)
r = put_row(ws, r, ["ZCRM_RECPTNO", "일자 2026.09.10 채번 이력",
                    "순번 1·2·3·4 순차 채번. 키 중복 없음"], BODY, OK_FILL)
r += 1
ws.cell(row=r, column=1, value=(
    "접수번호는 일자(8자리) + 순번(5자리) 형식이며, 두 프로그램이 같은 채번 테이블을 "
    "순차적으로 사용한다.")).font = BODY
widths(ws, {"A": 24, "B": 24, "C": 24, "D": 14})

ws = wb.create_sheet("9.사용데이터")
ws["A1"] = "테스트에 사용한 데이터"
ws["A1"].font = TITLE
ws["A2"] = "전 건 회수품 미접수 상태에서 선정. 접수번호가 없어 품질검사 데이터도 없던 건이다."
ws["A2"].font = NOTE
r = 4
r = put_row(ws, r, ["용도", "프로그램", "상담번호", "자재번호(부품)", "꼬리표NO",
                    "채번된 접수번호", "DB 변경"], H_FONT, H_FILL)
data_rows = [
    ("TC-07 적합수량 저장", "기존 ZCRM052", "2026070703501", "30036673A", "0010", "2026091000001", "있음"),
    ("TC-07 적합수량 저장", "신규 ZCRM052S", "2026062402790", "30030863A", "0093", "2026091000002", "있음"),
    ("TC-08 부적합수량 저장", "기존 ZCRM052", "2026070600189", "20050715A", "0201", "2026091000003", "있음"),
    ("TC-08 부적합수량 저장", "신규 ZCRM052S", "2026070202547", "30019272A", "0205", "2026091000004", "있음"),
    ("TC-09 검증 6종", "기존 ZCRM052", "2026070802937", "20070415A", "0205", "-", "없음 (전부 차단)"),
    ("TC-09 검증 6종", "신규 ZCRM052S", "2026070700456", "30033216A", "0215", "-", "없음 (전부 차단)"),
]
for d2 in data_rows:
    r = put_row(ws, r, list(d2), BODY, WARN_FILL if d2[6] == "있음" else OK_FILL)
r += 1
ws.cell(row=r, column=1, value="생성된 테스트 데이터 정리 기준").font = BOLD
r += 1
r = put_row(ws, r, ["테이블", "식별 조건"], H_FONT, H_FILL)
r = put_row(ws, r, ["ZCRM_RECPT", "RECPTDATE = 2026.09.10 또는 REMARK1 = 'ZTEST_20260909'"])
r = put_row(ws, r, ["ZCRM_RECPT_WEAR", "RECPTNO 2026091000001 ~ 2026091000004"])
r = put_row(ws, r, ["ZCRM_RECPTNO", "DATUM = 2026.09.10, SEQNO 1~4"])
r += 1
ws.cell(row=r, column=1, value=(
    "저장 4건의 접수부서명·접수자에는 ZTEST_20260909 가 들어 있다. 스크립트 상수를 "
    "오늘 날짜로 갱신하지 않은 채 실행한 결과다. 2026-09-09 에는 저장이 한 건도 성공하지 "
    "않았으므로 이 값을 가진 데이터는 2026-09-10 테스트로 생성된 4건뿐이다.")).font = NOTE
r += 2
ws.cell(row=r, column=1, value="부적합 사유 코드").font = BOLD
r += 1
r = put_row(ws, r, ["코드", "텍스트", "선정 사유"], H_FONT, H_FILL)
r = put_row(ws, r, ["Z02", "회수기간초과", "ZCRM_RECPT 실사용 빈도 최다"], BODY, OK_FILL)
r = put_row(ws, r, ["Z10", "해지대리점", "제외 - 해지점 자동 부적합 전용"])
widths(ws, {"A": 22, "B": 18, "C": 18, "D": 18, "E": 12, "F": 18, "G": 18})
ws.freeze_panes = "A5"
ws = wb.create_sheet("10.미해결·확인요청")
ws["A1"] = "미해결 항목과 확인 요청"
ws["A1"].font = TITLE
r = 3
ws.cell(row=r, column=1, value="2026-09-10 재테스트로 해소된 항목").font = BOLD
r += 1
r = put_row(ws, r, ["항목", "2026-09-09", "2026-09-10"], H_FONT, H_FILL)
r = put_row(ws, r, ["저장 경로 구독 검증",
                    "일반 데이터 저장 전부 거부. 저장 계열 비교 불가",
                    "제거됨. TC-07·08·09 전부 비교 완료"], BODY, OK_FILL)
r += 1
ws.cell(row=r, column=1, value="개발자 확인 요청").font = BOLD
r += 1
r = put_row(ws, r, ["#", "내용", "상태"], H_FONT, H_FILL)
r = put_row(ws, r, [1, "NCF_RSN_V (부적합사유 코드) 가 신규 필드카탈로그에 있는지, no_out 인지",
                    "2026-09-09 부터 미해결"], BODY, WARN_FILL)
r = put_row(ws, r, [2, "TC-09-06 메시지 끝 마침표 차이와 검증 시점 차이가 의도한 변경인지",
                    "2026-09-10 신규 발견"], BODY, WARN_FILL)
r += 1
ws.cell(row=r, column=1, value="데이터가 없어 미검증인 항목").font = BOLD
r += 1
r = put_row(ws, r, ["케이스", "필요 데이터", "확인 결과"], H_FONT, H_FILL)
r = put_row(ws, r, ["TC-12 품질검사 마감 잠금", "ZCRM_QRST_PRT-ZCLOSE = 'X'",
                    "접수번호 2026 전체 범위 0건"], BODY, NG_FILL)
r = put_row(ws, r, ["TC-13 해지점 자동 부적합", "AUFSD = 'Z8' 서비스점의 미접수 건",
                    "2026.04~07 4개 구간 대상 0건"], BODY, NG_FILL)
r = put_row(ws, r, ["TC-14 단납 보증기간", "제조번호 + ZCRM_WRTDG",
                    "마스터는 존재. 상담 데이터 연결 건 미확인"], BODY, NG_FILL)
r = put_row(ws, r, ["TC-10 책임사업장 F4", "-",
                    "ALV 셀 F4 팝업이 열리지 않음"], BODY, NG_FILL)
r = put_row(ws, r, ["TC-16 Header 변경 감지", "-", "미실행"], BODY, NG_FILL)
r += 2
ws.cell(row=r, column=1, value="테스트 종료 후 필수 조치").font = BOLD
r += 1
ws.cell(row=r, column=1, value="AND ZCRM_CICSVC~SUBS_TYPE = GC_SUBS_TYPE_S").font = Font(
    name="Consolas", bold=True, size=11, color="C00000")
r += 1
ws.cell(row=r, column=1, value=(
    "위 조건을 회수품 접수 SELECT 와 회수품 미접수 SELECT 양쪽 모두 원복한다. "
    "저장 경로의 구독 검증을 임시로 해제한 상태라면 그것도 함께 원복한다. "
    "원복 후 일반 상담번호가 신규에서 조회되지 않는지, 구독 데이터가 없을 때 0건으로 "
    "정상 동작하는지 확인한다.")).font = BODY
r += 2
ws.cell(row=r, column=1, value="생성된 테스트 데이터 4건은 9.사용데이터 시트의 정리 기준으로 식별한다.").font = BODY
widths(ws, {"A": 26, "B": 62, "C": 46})

ws = wb.create_sheet("11.화면캡처목록")
ws["A1"] = "화면 캡처 목록"
ws["A1"].font = TITLE
ws["A2"] = "경로: C:\\sap-test\\results\\shots\\  (PNG, 1929x1368)"
ws["A2"].font = NOTE
r = 4
r = put_row(ws, r, ["파일명", "케이스", "프로그램", "단계"], H_FONT, H_FILL)
shots = sorted(os.path.basename(p) for p in glob.glob(r"C:\sap-test\results\shots\*.png"))
for fn in shots:
    stem = fn[:-4]
    parts = stem.split("_")
    case = parts[0] if not stem.startswith("S10") else parts[1]
    prog = "신규" if ("NEW" in stem or "_S_" in stem or stem.endswith("_S")) else "기존"
    step = parts[-1]
    r = put_row(ws, r, [fn, case, prog, step])
widths(ws, {"A": 46, "B": 14, "C": 10, "D": 18})
ws.freeze_panes = "A5"
ws.cell(row=3, column=1, value="총 {}장".format(len(shots))).font = BOLD


# ---------------------------------------------------------------- 11. 주요 화면
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage

ws = wb.create_sheet("12.주요화면")
ws["A1"] = "주요 화면 캡처"
ws["A1"].font = TITLE
ws["A2"] = "원본: C:\sap-test\results\shots\ (1929x1368 PNG). 아래는 축소본."
ws["A2"].font = NOTE

PICKS = [
    ("TC01_OLD_2_result.png", "TC-01 센터조치 접수 조회 - 기존 ZCRM052 (5,312건)"),
    ("TC01_NEW_2_result.png", "TC-01 센터조치 접수 조회 - 신규 ZCRM052S (5,312건)"),
    ("TC01_OLD_1_selection.png", "선택화면 - 기존 (센터 필드 표시)"),
    ("TC01_NEW_1_selection.png", "선택화면 - 신규 (센터 필드 표시)"),
    ("TC03_OLD_1_selection.png", "미접수 선택화면 - 기존 (접수일자 사라짐)"),
    ("TC03_NEW_1_selection.png", "미접수 선택화면 - 신규 (접수일자 사라짐)"),
    ("S10_TC09-01_OLD_4_popup.png", "저장 확인 팝업 - 저장 하시겠습니까?"),
    ("S10_TC09-01_OLD_3_entered.png", "TC-09-01 값 입력 상태 - 기존"),
    ("S10_TC09-01_OLD_4b_after_confirm.png", "TC-09-01 수량 초과 검증 차단 - 기존"),
    ("S10_TC09-01_NEW_4b_after_confirm.png", "TC-09-01 수량 초과 검증 차단 - 신규"),
    ("TC15_ZCRMR052_onesel.png", "TC-15 상세조회 팝업 - 기존"),
    ("TC15_ZCRMR052S_onesel.png", "TC-15 상세조회 팝업 - 신규"),
    ("S10_TC07_OLD_4b_after_confirm.png", "TC-07 적합수량 저장 완료 - 기존"),
    ("S10_TC07_NEW_4b_after_confirm.png", "TC-07 적합수량 저장 완료 - 신규"),
    ("S10_TC08_OLD_4b_after_confirm.png", "TC-08 부적합수량 저장 완료 - 기존"),
    ("S10_TC08_NEW_4b_after_confirm.png", "TC-08 부적합수량 저장 완료 - 신규"),
    ("S10_TC09-06_OLD_4_after_save.png", "TC-09-06 입고일자 누락 차단 - 기존"),
    ("S10_TC09-06_NEW_4b_after_confirm.png", "TC-09-06 입고일자 누락 차단 - 신규"),
]

SHOTS = os.path.join(RAW, os.pardir, "shots")
THUMB = os.path.join(RAW, os.pardir, "shots", "_thumb")
os.makedirs(THUMB, exist_ok=True)

row = 4
placed = 0
for fn, caption in PICKS:
    src = os.path.join(SHOTS, fn)
    if not os.path.exists(src):
        continue
    dst = os.path.join(THUMB, fn)
    im = PILImage.open(src)
    im.thumbnail((900, 900))
    im.save(dst, optimize=True)
    ws.cell(row=row, column=1, value=caption).font = BOLD
    img = XLImage(dst)
    ws.add_image(img, "A{}".format(row + 1))
    rows_needed = int(img.height / 19) + 2
    row += rows_needed + 2
    placed += 1
ws.column_dimensions["A"].width = 130
print("images embedded:", placed)

for sheet in wb.worksheets:
    sheet.sheet_view.showGridLines = False

wb.save(OUT)
print("saved:", OUT)
print("sheets:", [s.title for s in wb.worksheets])
print("shots listed:", len(shots))
