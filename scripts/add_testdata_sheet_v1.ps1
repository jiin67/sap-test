# add_testdata_sheet_v1.ps1 : "테스트데이터" 시트 (SR03000035400 결과서 v1.00, 2026-10-02)
param([Parameter(Mandatory = $true)][string]$Path, [string]$SheetName = "테스트데이터", [int]$After = 7, [string]$Doc177 = "")
$ErrorActionPreference = "Stop"
$HDR = 2050429; $IN = 14804223; $OK = 14348258; $NG = 13551615
$d177 = $null
if ($Doc177 -and (Test-Path $Doc177)) { $d177 = Get-Content -Path $Doc177 -Raw -Encoding UTF8 | ConvertFrom-Json }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
$script:r = 1
try {
    $wb = $xl.Workbooks.Open($Path)
    foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq $SheetName) { $s.Delete() } }
    $ws = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($After)); $ws.Name = $SheetName
    function Title($t) { $c = $ws.Cells.Item($script:r, 1); $c.Value2 = $t; $c.Font.Size = 14; $c.Font.Bold = $true; $script:r += 2 }
    function Section($t) { $c = $ws.Cells.Item($script:r, 1); $c.Value2 = $t; $c.Font.Size = 11; $c.Font.Bold = $true; $c.Font.Color = $HDR; $script:r++ }
    function Head($cells) { for ($i = 0; $i -lt $cells.Count; $i++) { $c = $ws.Cells.Item($script:r, $i + 1); $c.Value2 = $cells[$i]; $c.Interior.Color = $HDR; $c.Font.Color = 16777215; $c.Font.Bold = $true; $c.HorizontalAlignment = -4108; $c.Borders.LineStyle = 1 }; $script:r++ }
    function Row($cells, $fill = $null) { for ($i = 0; $i -lt $cells.Count; $i++) { $c = $ws.Cells.Item($script:r, $i + 1); $c.NumberFormat = "@"; $c.Value2 = $cells[$i]; $c.Borders.LineStyle = 1; $c.VerticalAlignment = -4160; $c.WrapText = $true; if ($fill) { $c.Interior.Color = $fill } }; $script:r++ }
    function Note($t) { $c = $ws.Cells.Item($script:r, 1); $c.Value2 = $t; $c.Font.Italic = $true; $script:r += 2 }
    function Gap { $script:r++ }

    Title "전체 프로세스 테스트에 사용한 데이터 (2026-10-02, KDQ 210)"
    $ws.Cells.Item(2, 1).Value2 = "사용자 KKR003-1 / 기간월 2026.10 / 생성자 ZTEST1002E / 데이터와 생성 문서는 삭제하지 않고 남겨 둠. 생성 스크립트 scripts\mk1002_e2e.ps1 (-Set V -Part svc|new|req|m|sub|cnfrm)"
    $ws.Cells.Item(2, 1).Font.Italic = $true
    $script:r = 4

    Section "1. 따라하기 순서 (수동 재현용)"
    Head @("단계", "T-code / 작업", "입력값·조작", "확인")
    Row @("1", "SE16N 삽입 (ZCRM_CICSVC·ZCRM_CICSVC_NEW·ZCRM_REQ·ZCRM_CICCOST_M·_SUB·_ADD, 구독은 ZCRM_CHRG_CNFRM)", "상담번호·구독여부(SUBS_TYPE S)·보증잔여개월(SUBSWRTMONTH)·설치일·보증만료일·부품제조일·대리점(KUNNR_SVC)·센터(VKBUR)·조치일자 2026.10.02. CNFRM: SPMON 2026.10, CFDATE1 2026.10.02", "SE16N 으로 각 테이블 10행(CNFRM 7행)")
    Row @("2", "ZCRM177AS / ZCRM177A", "조치일자 2026.10.02~10.02 → 실행 → 행 체크 → 주문 → 납품 → 출고 (사용자 수행)", "구독 7건 / 일반 3건, ZCRM_CICCOST_M 오더·납품번호")
    Row @("3", "ZCRM052S / ZCRM052", "서비스점조치 + 회수품 미접수 + 상담번호 2026100290001~010 → 실행 → 적합수량 1(Z16 건은 부적합수량 1), 책임사업장 1100, 회수품입고일자 2026.10.02, 접수부서·접수자 → 저장 → 예", "구독 7건(Z16 2건) / 일반 3건, 금액 71,000 / 39,500")
    Row @("4", "ZCRM171S / ZCRM171", "정산 라디오 + 기간월 2026.10 → 실행 → 대리점 행 선택 → 본사확정 → 예 → 다시 선택 → 일괄회계전표생성 → 예 → 달력 오늘", "구독 3 대리점 / 일반 3 대리점, 전표 3100069966~971")
    Row @("5", "ZCRM058S", "전체 + 기간월 2026.10 + 센터(CT18/CT08/CT14/CT23) + 대리점 → 실행 → 선택 → 구독 무상 기술료/출장비 확정 → 예 → 다시 선택 → 일괄회계전표생성 → 예 → 달력 오늘", "전표 3100069972~975")
    Row @("6", "ZCRM058A", "전체 + 기간월 2026.10 + 센터 + 대리점 → 실행 → 선택 → 무상지원금확정 → 예 → 다시 선택 → 대리점 확정 → 예 → 달력 오늘 → 다시 선택 → 일괄회계전표생성 → 예 → 달력 오늘", "전표 3100069976~978")
    Row @("7", "ZCRM060S / ZCRM060A", "기간월 2026.10 + 센터 + 대리점 → 실행", "058 과 같은 건수·금액")
    Note "※ 058/060 선택화면의 센터는 단일값(범위 없음)이라 센터별로 조회한다. 058S 는 CIC 확정일(ZCRM_CHRG_CNFRM)이 없으면 전표가 막힌다. 171·058 모두 버튼을 누르면 선택이 풀리므로 다음 버튼 전에 행을 다시 선택한다."

    Section "2. 생성한 상담건 (상담번호 2026100290001~010, 조치일자 2026.10.02, 부품 30031356A 룸콘 NR-63S)"
    Head @("상담번호", "대리점 / 센터", "조건", "구독여부·잔여개월", "설치일 → 보증만료일", "부품제조일", "접수번호 / 금액", "058 지원금")
    Row @("2026100290001", "100187 하나수도 / CT18", "A 구독 N년이후", "S / 12", "2023.10.01 → 2026.10.01 (36개월 3년)", "2025.12.13", "2026100200001 / 71,000 (ZPA2)", "구독 5,000") $OK
    Row @("2026100290002", "100187 / CT18", "B 구독 N년이내", "S / 0", "2024.10.01 → 2026.10.01 (24개월 2년)", "2025.12.13", "2026100200002 / 39,500 (PR00)", "구독 5,000") $OK
    Row @("2026100290003", "100187 / CT18", "C 구독 기간만료", "S / 12", "2023.10.01 → 2026.10.01 (3년)", "2020.01.01", "2026100200003 / 부적합 Z16", "구독 5,000") $IN
    Row @("2026100290004", "100187 / CT18", "D 일반", "공백 / 0", "-", "2025.12.13", "2026100200008 / 39,500", "일반 5,000") $OK
    Row @("2026100290005", "100216 경진상사 / CT08", "A 구독 N년이후", "S / 24", "2021.10.01 → 2026.10.01 (60개월 5년)", "2025.12.13", "2026100200004 / 71,000", "구독 5,000") $OK
    Row @("2026100290006", "100216 / CT08", "D 일반", "공백 / 0", "-", "2025.12.13", "2026100200009 / 39,500", "일반 5,000") $OK
    Row @("2026100290007", "100241 통영대리점 / CT14", "A 구독 N년이후", "S / 12", "2022.10.01 → 2026.10.01 (48개월 4년)", "2025.12.13", "2026100200005 / 71,000", "구독 5,000") $OK
    Row @("2026100290008", "100241 / CT14", "B 구독 N년이내", "S / 0", "2025.10.01 → 2026.10.01 (12개월 1년)", "2025.12.13", "2026100200006 / 39,500", "구독 5,000") $OK
    Row @("2026100290009", "100282 보성대리점 / CT23", "C 구독 기간만료", "S / 0", "2024.10.01 → 2026.10.01 (2년)", "2020.01.01", "2026100200007 / 부적합 Z16", "구독 5,000") $IN
    Row @("2026100290010", "100282 / CT23", "D 일반", "공백 / 0", "-", "2025.12.13", "2026100200010 / 39,500", "일반 5,000") $OK
    Note "※ 회수가능제조일 = 설치일 기준으로 N년이후/이내에 따라 계산. 부품제조일(2020.01.01)이 그보다 앞서면 구독 모드에서 Z16 자동 부적합. 대리점가 PR00 39,500 / 소비자가 ZPA2 71,000 / 회수품 지원금은 + VAT 10%."

    Section "3. 생성된 전표"
    Head @("구분", "대리점", "전표", "금액", "계정 / 코스트센터 / 예산센터", "상태")
    Row @("④ 회수품정산 구독 (ZCRM171S, S08)", "100187", "3100069966", "121,550", "5328210 / 260014 / 240074", "정상") $OK
    Row @("④ 회수품정산 구독", "100216", "3100069967", "78,100", "5328210 / 260014 / 240074", "정상") $OK
    Row @("④ 회수품정산 구독", "100241", "3100069968", "121,550", "5328210 / 260038 / 240074", "정상") $OK
    Row @("④ 회수품정산 일반 (ZCRM171, Z08)", "100187 / 100216 / 100282", "3100069969 / 970 / 971", "각 43,450", "5328210 / 260014·260014·260010 / 260002", "정상") $OK
    Row @("⑤ 무상지원금 구독 (ZCRM058S, S01)", "100187 / 100216 / 100241 / 100282", "3100069972 / 973 / 974 / 975", "15,000 / 5,000 / 10,000 / 5,000", "5328170 / 260014·260014·260038·260010 / 240074", "정상 (예산 오류 없음)") $OK
    Row @("⑤ 무상지원금 일반 (ZCRM058A, Z01)", "100187 / 100216 / 100282", "3100069976 / 977 / 978", "각 5,000", "5328170 / 260014·260014·260010 / 260002", "정상") $OK
    if ($d177) { foreach ($x in $d177.rows) { Row @($x[0], $x[1], $x[2], $x[3], $x[4], $x[5]) $OK } }
    else { Row @("② 무상출고 (ZCRM177AS/177A)", "-", "(사용자 수행 대기)", "-", "5328130 / 출고 센터 / 공백 (현행)", "미실행") $IN }
    Gap

    Section "4. 단계별 조회 결과 요약"
    Head @("단계", "구독 화면", "일반 화면", "확인 내용")
    Row @("② 센터매출 (177)", "7건", "3건", "상담번호 단위 분리, 혼재 없음") $OK
    Row @("③ 회수품접수 (052)", "7건 (Z16 2건)", "3건", "구독 모드에서만 소비자가·Z16 적용") $OK
    Row @("④ 회수품정산 (171)", "3 대리점 (100282 는 부적합만이라 제외)", "3 대리점", "대리점 + 구독구분 단위 집계·전표") $OK
    Row @("⑤ 무상지원금 (058)", "4 대리점 35,000", "3 대리점 15,000", "같은 대리점이 구독·일반으로 나뉘어 산정, S 행 CFDATE1 은 CIC 확정일 표시") $OK
    Row @("⑥ 지원금상세 (060)", "4 대리점 (기술료·출장비 컬럼)", "3 대리점", "058 과 일치") $OK

    $ws.Columns.Item(1).ColumnWidth = 22
    $ws.Columns.Item(2).ColumnWidth = 30
    $ws.Columns.Item(3).ColumnWidth = 30
    $ws.Columns.Item(4).ColumnWidth = 22
    $ws.Columns.Item(5).ColumnWidth = 34
    $ws.Columns.Item(6).ColumnWidth = 14
    $ws.Columns.Item(7).ColumnWidth = 26
    $ws.Columns.Item(8).ColumnWidth = 14
    $ws.Rows("1:$script:r").EntireRow.AutoFit() | Out-Null
    $ws.Activate(); $xl.ActiveWindow.DisplayGridlines = $false
    $wb.Worksheets.Item(1).Activate()
    $wb.Save(); $wb.Close($false)
    "sheet '$SheetName' added"
}
finally { $xl.ScreenUpdating = $true; $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); [GC]::Collect() }
