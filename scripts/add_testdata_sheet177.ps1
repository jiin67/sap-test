# add_testdata_sheet177.ps1 : "테스트데이터" sheet for the 2026-09-19 ZCRMR177_1 (ZCRM177A / ZCRM177AS) test.
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [string]$SheetName = "테스트데이터"
)
$ErrorActionPreference = "Stop"
$HDR = 2050429; $IN = 14804223; $OK = 14348258; $NG = 13551615

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
$script:r = 1
try {
    $wb = $xl.Workbooks.Open($Path)
    foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq $SheetName) { $s.Delete() } }
    $ws = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item(6))
    $ws.Name = $SheetName

    function Title($t) { $c = $ws.Cells.Item($script:r, 1); $c.Value2 = $t; $c.Font.Size = 14; $c.Font.Bold = $true; $script:r += 2 }
    function Section($t) { $c = $ws.Cells.Item($script:r, 1); $c.Value2 = $t; $c.Font.Size = 11; $c.Font.Bold = $true; $c.Font.Color = $HDR; $script:r++ }
    function Head($cells) { for ($i = 0; $i -lt $cells.Count; $i++) { $c = $ws.Cells.Item($script:r, $i + 1); $c.Value2 = $cells[$i]; $c.Interior.Color = $HDR; $c.Font.Color = 16777215; $c.Font.Bold = $true; $c.HorizontalAlignment = -4108; $c.Borders.LineStyle = 1 }; $script:r++ }
    function Row($cells, $fill = $null) { for ($i = 0; $i -lt $cells.Count; $i++) { $c = $ws.Cells.Item($script:r, $i + 1); $c.NumberFormat = "@"; $c.Value2 = $cells[$i]; $c.Borders.LineStyle = 1; $c.VerticalAlignment = -4160; $c.WrapText = $true; if ($fill) { $c.Interior.Color = $fill } }; $script:r++ }
    function Gap { $script:r++ }
    function Note($t) { $c = $ws.Cells.Item($script:r, 1); $c.Value2 = $t; $c.Font.Italic = $true; $script:r += 2 }

    Title "테스트에 사용한 데이터 (2026-09-19 ZCRMR177_1 / ZCRM177A · ZCRM177AS)"
    $ws.Cells.Item(2, 1).Value2 = "KDQ 210 / 사용자 KKR003-1 / 테스트 데이터는 사용자 방침에 따라 삭제하지 않고 남겨 둠. 문서(오더·납품·자재문서)는 생성되지 않았음"
    $ws.Cells.Item(2, 1).Font.Italic = $true

    Section "0. 따라하기 순서"
    Head @("단계", "내용")
    Row @("1", "SE16N 으로 ZCRM_CICSVC 에 상담건 삽입. 복사 원본은 2026091890001(일반) / 2026091890002(구독 S). CSLTNO·CSLTNO_ITEM·TAGNO·SUBS_TYPE·SUBSWRTMONTH·WRTDATE·이름/주소(ZTEST_20260919)만 바꾼다. CLMART='A' 필수(조인 조건)")
    Row @("2", "SE16N 으로 ZCRM_CICCOST_M 에 같은 상담번호 행 삽입. 복사 원본 2026091891002. VKORG=ASKR, CENTER_CD=CT04, TECHNAME=ZTEST, INDATE(조치일자)=2026.09.19, ISFREE(유상 공백 / 무상 'Y')")
    Row @("3", "SE16N 으로 ZCRM_CICCOST_SUB 에 COST=0 인 무상부품 행 삽입(상담번호당 1행). 이 행이 없으면 조회에서 상담번호가 통째로 빠진다 - ZCRMR177_1F01 165행 '올유상인건 지운다'")
    Row @("4", "ZCRM177A / ZCRM177AS 실행 → 전기일자 오늘, 회사코드 1100, 판매조직 ASKR, 조치일자 2026.09.19 ~ 2026.09.19, '완료건 포함하여 조회' 체크 → 실행(F8)")
    Row @("5", "결과 ALV 에서 '구독여부' 컬럼이 보이지 않으면 레이아웃 변경(Ctrl+F8 또는 레이아웃 선택 → 레이아웃 변경)에서 '숨긴 열'의 '구독여부'를 '표시된 열'로 옮긴다 (결함 D-01)")
    Note "※ 문서처리(주문생성·납품생성·출고전기)는 ALV 왼쪽 '선택' 체크박스를 켠 뒤 상단 버튼을 누른다. 이 체크박스는 GUI 스크립팅으로 전달되지 않아 이번 테스트에서는 수동 확인이 필요하다."

    Section "1. 생성한 테스트 데이터 (상담번호 2026091990001~013, 생성자 ZTEST0919)"
    Head @("케이스", "상담번호", "SUBS_TYPE", "SUBSWRTMONTH", "WRTDATE", "ISFREE", "조치일자", "조회되는 화면 / 구독여부")
    Row @("A 일반 유상", "2026091990001", "(공백)", "0", "(공백)", "(공백)", "2026.09.19", "ZCRM177A / 유상") $IN
    Row @("B 일반 무상", "2026091990002", "(공백)", "0", "(공백)", "Y", "2026.09.19", "ZCRM177A / 무상") $IN
    Row @("C 구독 N년이후", "2026091990003", "S", "12", "2027.09.19", "Y", "2026.09.19", "ZCRM177AS / 구독무상(N년이후)") $IN
    Row @("D 구독 음수(시도)", "2026091990004", "S", "0 (-6 저장 실패)", "2026.03.19", "Y", "2026.09.19", "ZCRM177AS / 구독무상(N년이내)") $NG
    Row @("E 구독 0 경계", "2026091990005", "S", "0", "2026.09.19", "Y", "2026.09.19", "ZCRM177AS / 구독무상(N년이내)") $IN
    Row @("F 보증일 공란", "2026091990006", "S", "3", "(공백)", "Y", "2026.09.19", "ZCRM177AS / 구독무상(N년이후)") $IN
    Row @("G 다품목 혼재", "2026091990007", "품목1 S / 품목2 (공백)", "품목1 12", "품목1 2027.09.19", "Y", "2026.09.19", "양쪽 모두 표시 (결함 D-02)") $NG
    Row @("H 다품목 동일", "2026091990008", "품목1·2 모두 S", "12", "2027.09.19", "Y", "2026.09.19", "ZCRM177AS / 1행만") $IN
    Row @("I 기간 시작일", "2026091990009", "S", "12", "2027.09.19", "Y", "2026.09.17", "기간 09.17~19 조회 시 포함") $IN
    Row @("J 기간 밖", "2026091990010", "S", "12", "2027.09.19", "Y", "2026.09.16", "기간 09.17~19 조회 시 제외") $IN
    Row @("L 문서처리용 구독", "2026091990012", "S", "12", "2027.09.19", "Y", "2026.09.19", "ZCRM177AS / 구독무상(N년이후)") $IN
    Row @("M 문서처리용 일반", "2026091990013", "(공백)", "0", "(공백)", "Y", "2026.09.19", "ZCRM177A / 무상") $IN
    Note "※ 공통값: ZCRM_CICSVC CLMART='A', PRICETYPE='N', RESULTORG='B', 자재 PNCB3540026K001, 부품 30031356A. ZCRM_CICCOST_M VKORG=ASKR, CENTER_CD=CT04, TECHNAME=ZTEST, ITEMCD=NCB3540026K, CHGITEMCD1=30031356A. ZCRM_CICCOST_SUB ITEM 10 / 30031356A / QTY 1 / COST 0. 문서처리용 2건(012·013)에는 ZCRM_CICCOST_ADD 행도 있다."

    Section "2. 만들지 못한 데이터 (제약 사항)"
    Head @("항목", "시도", "결과", "영향")
    Row @("SUBSWRTMONTH 음수", "SE16N 삽입 '-6', SE16N 편집 '-6' 및 '6-'", "3회 모두 오류 메시지 없이 저장되지 않음(변경 완료 팝업 없음). ZCRM_CICSVC 는 SM30 유지보수 다이얼로그도 없음", "TC-10 · TC-17 미실행. 소스 분기는 '>0 이 아니면 이내' 이므로 0 케이스(TC-18)로 같은 경로 확인") $NG
    Row @("SUBS_TYPE 'S' 외 값", "SE16N 삽입 'X'", "도메인이 거부해 저장되지 않음", "TC-42 의 'S 외 값' 은 SAP 쪽에서 존재할 수 없음 - 값 영역은 공백/S 뿐") $NG
    Gap

    Section "3. 조회 결과 요약 (조치일자 2026.09.19 ~ 2026.09.19, 완료건 포함)"
    Head @("화면", "건수", "상담번호", "비고")
    Row @("ZCRM177A (일반)", "4", "2026091990001, 002, 007, 013", "구독여부 유상 1 / 무상 3") $OK
    Row @("ZCRM177AS (구독)", "7", "2026091990003, 004, 005, 006, 007, 008, 012", "구독무상(N년이후) 5 / (N년이내) 2") $OK
    Row @("대상 전체(원천)", "10", "2026091990001~008, 012, 013", "4 + 7 = 11 ≠ 10 → 2026091990007 이 양쪽에 중복 (결함 D-02)") $NG
    Gap

    Section "4. 대량 조회 (조치일자 2026.01.01 ~ 2026.09.19, 완료건 포함)"
    Head @("화면", "건수", "소요시간", "비고")
    Row @("ZCRM177A", "21,575", "약 73초", "TIME_OUT·덤프 없음") $OK
    Row @("ZCRM177AS", "13", "약 19초", "기존 09-18 테스트 데이터 2건 포함") $OK
    Gap

    Section "5. 수정 전/후 구분 기준 (소스 주석)"
    Head @("항목", "내용")
    Row @("주석 표식", "**<<SR03000035400 보일러 서비스 구독 대리점 비용 정산 건 … **>>SR03000035400 보일러 서비스 구독 대리점 비용 정산 건")
    Row @("블록 수", "13곳 - ZCRMR177_1 1, ZCRMR177_1T01 2, ZCRMR177_1C01 1, ZCRMR177_1F01 8, ZCRMR177_1O01 1")
    Row @("주요 변경", "INITIALIZATION 에 PERFORM zf_set_subs_mode / GT_DATA 에 SUBS_TYPE·SUBSWRTMONTH·SUBS_TEXT 추가 / GC_TCODE_SUBS='ZCRM177AS', GR_SUBS 레인지 / SELECT 에 c~subs_type·c~subswrtmonth 와 조건 c~subs_type IN gr_subs / 행마다 ZF_SET_SUBS_TEXT / 필드카탈로그 SUBS_TEXT / GT_AUTO-PROGRAM_NM 분기 3곳 / SET TITLEBAR 'ZSUB'")
    Row @("전체 소스", "RS_ABAP_SOURCE_SCAN 으로 활성 소스 8,923행(INCLUDE 11개) 덤프 → results\src\ZCRMR177_1_scan.txt, 수정 블록 정리본 → results\src\ZCRMR177_1_diffblocks.txt")

    $ws.Columns.Item(1).ColumnWidth = 22
    $ws.Columns.Item(2).ColumnWidth = 30
    $w = @(22, 30, 22, 20, 16, 12, 14, 34)
    for ($i = 0; $i -lt $w.Count; $i++) { $ws.Columns.Item($i + 1).ColumnWidth = $w[$i] }
    $ws.Rows("1:$script:r").EntireRow.AutoFit() | Out-Null
    $ws.Activate()
    $xl.ActiveWindow.DisplayGridlines = $false
    $wb.Worksheets.Item(1).Activate()
    $wb.Save()
    $wb.Close($false)
    "sheet '$SheetName' added to $Path"
}
finally {
    $xl.ScreenUpdating = $true
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
