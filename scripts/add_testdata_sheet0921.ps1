# add_testdata_sheet0921.ps1 : "테스트데이터" sheet for the 2026-09-21~22 추가 개선 / 회귀 결과서
param([Parameter(Mandatory = $true)][string]$Path, [string]$SheetName = "테스트데이터")
$ErrorActionPreference = "Stop"
$HDR = 2050429; $IN = 14804223; $OK = 14348258; $NG = 13551615
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
$script:r = 1
try {
    $wb = $xl.Workbooks.Open($Path)
    foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq $SheetName) { $s.Delete() } }
    $ws = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item(6)); $ws.Name = $SheetName
    function Title($t) { $c = $ws.Cells.Item($script:r, 1); $c.Value2 = $t; $c.Font.Size = 14; $c.Font.Bold = $true; $script:r += 2 }
    function Section($t) { $c = $ws.Cells.Item($script:r, 1); $c.Value2 = $t; $c.Font.Size = 11; $c.Font.Bold = $true; $c.Font.Color = $HDR; $script:r++ }
    function Head($cells) { for ($i = 0; $i -lt $cells.Count; $i++) { $c = $ws.Cells.Item($script:r, $i + 1); $c.Value2 = $cells[$i]; $c.Interior.Color = $HDR; $c.Font.Color = 16777215; $c.Font.Bold = $true; $c.HorizontalAlignment = -4108; $c.Borders.LineStyle = 1 }; $script:r++ }
    function Row($cells, $fill = $null) { for ($i = 0; $i -lt $cells.Count; $i++) { $c = $ws.Cells.Item($script:r, $i + 1); $c.NumberFormat = "@"; $c.Value2 = $cells[$i]; $c.Borders.LineStyle = 1; $c.VerticalAlignment = -4160; $c.WrapText = $true; if ($fill) { $c.Interior.Color = $fill } }; $script:r++ }
    function Note($t) { $c = $ws.Cells.Item($script:r, 1); $c.Value2 = $t; $c.Font.Italic = $true; $script:r += 2 }
    function Gap { $script:r++ }

    Title "테스트에 사용한 데이터 (2026-09-21~22, KDQ 210)"
    $ws.Cells.Item(2, 1).Value2 = "사용자 KKR003-1 / 기간월 2026.09 / 데이터와 생성 문서는 사용자 방침에 따라 삭제하지 않고 남겨 둠"
    $ws.Cells.Item(2, 1).Font.Italic = $true

    Section "1. 생성한 상담건 (생성자 ZTEST0921, 조치일자 2026.09.21, 부품 30031356A)"
    Head @("상담번호", "대리점 / 센터", "구독구분", "SUBSWRTMONTH", "부품제조일", "용도", "결과")
    Row @("2026092190001", "100034 / 0520", "S", "12 (N년이후)", "2025.12.13", "최초 금액 확인(반영 전)", "접수 2026092100001, 39,500 (ZPA2 반영 전 저장)") $IN
    Row @("2026092190002", "100034 / 0520", "S", "12", "2020.01.01", "Z02 자동판정", "Z02 회수기간초과 / 부적합 (미저장)") $OK
    Row @("2026092190003", "100050 / CT03", "(공백)", "0", "2025.12.13", "비구독 회귀", "접수 2026092100002, 39,500") $OK
    Row @("2026092190004", "100034 / 0520", "S", "0 (N년이내)", "2025.12.13", "Z02 자동판정(N년이내)", "Z02 회수기간초과 / 부적합 (미저장)") $OK
    Row @("2026092190005", "100034 / 0520", "S", "12 (N년이후)", "2025.12.13", "소비자가(ZPA2) 검증", "접수 2026092100003, 71,000") $OK
    Row @("2026092190006", "100034 / 0520", "S", "0 (N년이내)", "2026.08.01", "대리점가(PR00) 검증", "접수 2026092100004, 39,500") $OK
    Note "※ 단가: PR00 대리점가 39,500원(KNUMH 0009916890) / ZPA2 소비자가 71,000원(KNUMH 0009922402). 센터 0520 = 구독사업팀(TVKBT)."

    Section "2. 생성·취소된 FI 전표 (회사코드 1100 / 2026 / KR)"
    Head @("전표번호", "대리점 / 구분", "비용계정", "코스트센터", "금액", "상태")
    Row @("3100069939", "100034 구독", "5328210 A/S부품비 (분기 반영 전)", "260005", "43,450", "역분개 9930001226") $NG
    Row @("3100069940", "100050 일반", "5328210", "260011", "86,900", "유효 (분기 반영 전 일반 기준값)") $OK
    Row @("3100069941", "100034 구독", "5328170 A/S지원금 (Z11 임시값)", "260005", "165,000", "역분개 9930001227") $NG
    Row @("3100069942", "100090 일반", "5328210", "260013", "43,450", "유효 (분기 반영 후 일반 회귀)") $OK
    Row @("3100069943", "100034 구독", "4305100 용역매출원가", "260005", "165,000", "유효 (최종)") $OK
    Gap

    Section "3. 설정 테이블 변경 (사용자 등록, 2026-09-22 확인)"
    Head @("테이블", "키", "내용")
    Row @("ZCRM_FIPOST_T (뷰 ZCRM_FIPOST_V)", "Z11 / 1100 / 1100", "HKONT 2103100, NEWKO 4305100 용역매출원가, 전기키 31/40, 세금 A1, SGTXT1 'test_용역매출원가'")
    Row @("ZCRM_KOSTL_T (AS정산 코스트센터 유지보수뷰)", "Z11 / 1100 / 0520·CT01~CT24", "Z08 과 같은 코스트센터 복사, 0520 = 240074")
    Row @("정산유형 텍스트", "Z11", "'test_기술비+출장비' (운영 전 정리 필요 - OPEN_ITEMS O-4)")
    Gap

    Section "4. 따라하기 순서"
    Head @("단계", "T-code", "입력")
    Row @("1", "SE16N", "ZCRM_CICSVC 에 상담건 삽입(VKBUR 0520, SUBS_TYPE S, SUBSWRTMONTH, PRDDATE_PRT) + ZCRM_REQ 삽입")
    Row @("2", "ZCRM052S", "서비스점조치 + 회수품 미접수 + 상담번호 → 적합수량 1 / 책임사업장 1100 / 회수품입고일자 → 저장 → '예'. Z02 건은 조회만 해도 부적합사유가 붙어 있음")
    Row @("3", "SE16N ZCRM_RECPT", "NETWR 확인: N년이후 71,000 / N년이내·비구독 39,500")
    Row @("4", "ZCRM171S", "정산 + 기간월 2026.09 + 센터 0520 → 행 선택 → 본사 확정 → 다시 선택 → 일괄회계전표생성 → 달력 '오늘'")
    Row @("5", "SE16N BSEG", "전표번호로 조회 → 40 항목 HKONT 4305100 확인")
    Note "※ 버튼을 누를 때마다 행 선택이 풀리므로 다음 버튼 전에 다시 선택한다. 전표취소 버튼 이름은 '일괄회계전표삭제'."

    $w = @(18, 22, 14, 16, 14, 26, 36); for ($i = 0; $i -lt $w.Count; $i++) { $ws.Columns.Item($i + 1).ColumnWidth = $w[$i] }
    $ws.Rows("1:$script:r").EntireRow.AutoFit() | Out-Null
    $ws.Activate(); $xl.ActiveWindow.DisplayGridlines = $false
    $wb.Worksheets.Item(1).Activate(); $wb.Save(); $wb.Close($false)
    "sheet '$SheetName' added"
}
finally { $xl.ScreenUpdating = $true; $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
