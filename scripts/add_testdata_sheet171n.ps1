# add_testdata_sheet194.ps1 : "테스트데이터" sheet for the 2026-09-14 ZCRMR194 → ZCRM052 test (v0.01).
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [string]$SheetName = "테스트데이터_v0.03"
)
$ErrorActionPreference = "Stop"
$HDR = 2050429; $IN = 14804223; $OK = 14348258; $NG = 13551615

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false
$script:r = 1
try {
    $wb = $xl.Workbooks.Open($Path)
    foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq $SheetName) { $s.Delete() } }
    $ws = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item(7))
    $ws.Name = $SheetName

    function Title($t) { $c = $ws.Cells.Item($script:r, 1); $c.Value2 = $t; $c.Font.Size = 14; $c.Font.Bold = $true; $script:r += 2 }
    function Section($t) { $c = $ws.Cells.Item($script:r, 1); $c.Value2 = $t; $c.Font.Size = 11; $c.Font.Bold = $true; $c.Font.Color = $HDR; $script:r++ }
    function Head($cells) { for ($i = 0; $i -lt $cells.Count; $i++) { $c = $ws.Cells.Item($script:r, $i + 1); $c.Value2 = $cells[$i]; $c.Interior.Color = $HDR; $c.Font.Color = 16777215; $c.Font.Bold = $true; $c.HorizontalAlignment = -4108; $c.Borders.LineStyle = 1 }; $script:r++ }
    function Row($cells, $fill = $null) { for ($i = 0; $i -lt $cells.Count; $i++) { $c = $ws.Cells.Item($script:r, $i + 1); $c.NumberFormat = "@"; $c.Value2 = $cells[$i]; $c.Borders.LineStyle = 1; $c.VerticalAlignment = -4160; $c.WrapText = $true; if ($fill) { $c.Interior.Color = $fill } }; $script:r++ }
    function Gap { $script:r++ }
    function Note($t) { $c = $ws.Cells.Item($script:r, 1); $c.Value2 = $t; $c.Font.Italic = $true; $script:r += 2 }

    Title "테스트 데이터 v0.03 (2026-09-18 23시, 2차 데이터 전체 재수행)"
    $ws.Cells.Item(2, 1).Value2 = "KDQ 210 / 기간월 2026.09 / 생성자 ZTEST0918C, 고객명 ZTEST_20260918C / 데이터는 삭제하지 않고 남겨 둠. 만드는 방법은 '테스트데이터' 시트 0번과 동일"
    $ws.Cells.Item(2, 1).Font.Italic = $true

    Section "1. 2차 테스트 데이터 (요청번호 2026091892001)"
    Head @("상담번호-품목", "대리점 / 센터", "구분", "꼬리표", "접수번호", "용도", "종료 시점 상태")
    Row @("2026091892001-1", "100041 남양주경동보일러 / CT01", "구독 S", "9301", "2026091800018", "혼재 대리점(G) 구독 쪽: TC-01·04~07·14·15·17·20~22·28·31·35~43", "구독 노랑") $IN
    Row @("2026091892002-1", "100041 남양주경동보일러 / CT01", "일반", "9302", "2026091800015", "혼재 대리점(G) 일반 쪽: TC-02·18·29·35~39", "일반 노랑") $IN
    Row @("2026091892003-1, -2", "100043 경동보일러의정부점 / CT01", "구독 S", "9303, 9304", "2026091800019, 020", "다품목(H): TC-03·11·19·25·27·30·33", "구독 노랑") $IN
    Row @("2026091892004-1", "100044 의정부경동보일러 / CT01", "구독 S", "9305", "2026091800021", "TC-03·16·23·24·30, X-01·02", "구독 노랑") $IN
    Row @("2026091892005-1", "100045 OK상사 / CT01", "구독 S", "9306", "2026091800022", "재검증 대상: TC-03·12·19·27·30·33, X-01·02", "구독 노랑") $IN
    Row @("2026091892006-1", "100047 (주)백일 / CT17", "구독 S", "9307", "2026091800023", "재검증 대상(I): TC-08", "구독 빨강") $IN
    Row @("2026091892007-1", "100047 (주)백일 / CT17", "일반", "9308", "2026091800016", "재검증 대상(I): TC-09", "일반 빨강") $IN
    Row @("2026091892008-1", "100050 경동나비엔계양대리점 / CT03", "일반", "9309", "2026091800017", "일반 전용: TC-13·26·44", "일반 빨강") $IN
    Note "※ 모두 해지점이 아닌 정상 대리점(KNVV 오더보류 없음). 부품 30031356A, 부품가 39,500 → 정산 43,450(VAT 3,950), 100043 은 2건 86,900(7,900)."

    Section "2. v0.03 에서 생성된 FI 전표 (회사코드 1100 / 2026, 전표유형 KR)"
    Head @("전표번호", "대리점 / 구분", "케이스", "금액 / 지급조건", "취소(역분개) 전표", "현재")
    Row @("3100069931", "100044 구독", "TC-24", "43,450 / E030", "9930001222 (TC-30)", "취소됨")
    Row @("3100069932", "100041 구독", "TC-17", "43,450 / E000", "9930001219 (TC-28)", "취소됨")
    Row @("3100069933", "100041 일반", "TC-18·38", "43,450 / E000", "9930001220 (TC-29)", "취소됨")
    Row @("3100069934", "100043 구독", "TC-19", "86,900 / E000", "9930001221 (TC-30)", "취소됨")
    Row @("3100069935", "100045 구독", "TC-19", "43,450 / E000", "9930001223 (TC-30)", "취소됨")
    Row @("3100069936", "100050 일반", "TC-44", "43,450 / E000", "9930001224", "취소됨")
    Note "※ 재검증 차단 케이스(TC-25·26·27·33)에서는 전표·역분개가 생기지 않았다."

    Section "3. 재검증 케이스에서 바꾼 SUBS_TYPE (모두 원복 완료)"
    Head @("케이스", "상담번호-품목", "변경", "결과")
    Row @("TC-11", "2026091892003-2", "S → 공백 → S", "본사확정 차단") $OK
    Row @("TC-08", "2026091892006-1", "S → 공백 → S", "본사확정 차단") $OK
    Row @("TC-09", "2026091892007-1", "공백 → S → 공백", "본사확정 차단 (ZCRM171)") $OK
    Row @("X-01 / X-02 / TC-27", "2026091892005-1", "S → 공백 → S (3회)", "다중 확정·확정취소·전표생성 전체 차단") $OK
    Row @("TC-16", "2026091892004-1", "S → 공백 → S", "확정취소 차단") $OK
    Row @("TC-25 / TC-33", "2026091892003-1", "S → 공백 → S (2회)", "전표생성·전표취소 차단") $OK
    Row @("TC-26", "2026091892008-1", "공백 → S → 공백", "전표생성 차단 (ZCRM171)") $OK
    Gap

    Section "4. 종료 시점 ZCRM_RECPTF (정산기간 2026.09) - 2차 데이터분"
    Head @("대리점", "구독여부", "정산대상", "본사확정일", "회계전표번호")
    Row @("0000100041", "(공백)", "1", "2026.09.18", "")
    Row @("0000100041", "S", "1", "2026.09.18", "")
    Row @("0000100043", "S", "2", "2026.09.18", "")
    Row @("0000100044", "S", "1", "2026.09.18", "")
    Row @("0000100045", "S", "1", "2026.09.18", "")
    Note "※ 1차 데이터분(100032·100035·100036·100037·100090·100273)은 '테스트데이터' 시트 참조. 100036 S 행의 취소된 전표번호 3100069924 는 여전히 남아 있음(개발자 정리 필요)."

    $w = @(24, 34, 22, 26, 34, 40, 22)
    for ($i = 0; $i -lt $w.Count; $i++) { $ws.Columns.Item($i + 1).ColumnWidth = $w[$i] }
    $ws.Activate(); $xl.ActiveWindow.DisplayGridlines = $false
    $wb.Worksheets.Item(1).Activate()
    $wb.Save()
    "sheet added: $SheetName rows=$($script:r)"
    $wb.Close($true)
}
finally {
    $xl.ScreenUpdating = $true; $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); [GC]::Collect()
}
