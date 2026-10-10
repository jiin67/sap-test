# add_testdata_sheet4.ps1 : "테스트데이터" sheet for the 2026-09-13 통합 테스트 (v0.04).
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

    Title "테스트에 사용한 데이터 (2026-09-13 통합 테스트)"
    $ws.Cells.Item(2, 1).Value2 = "KDQ 210 / 기간월 2026.07 / 모든 데이터는 테스트 후 원복 완료 (역전표 이력은 FI 에 남음)"
    $ws.Cells.Item(2, 1).Font.Italic = $true

    Section "0. 데이터 준비 방식"
    Head @("항목", "내용")
    Row @("구독 구분", "ZCRM_CICSVC_NEW-SUBS_TYPE 4건을 공백 → 'S' 로 변경 (SE16N), 테스트 후 공백 복원")
    Row @("기술료/출장비", "QA 의 무상·구독 상담건에는 ZCRM_CICCOST_M 행이 없어(실제 구독건 포함), 사용자 승인 하에 테스트 행 4건을 SE16N 으로 삽입 (생성자 ZTEST0913). 테스트 후 4건 삭제")
    Row @("대리점확정일", "RFC ZCRMF03_RFC_ZCRM_CHRG_CNFRM 으로 등록(S)·취소(C). 테스트 후 0건")
    Row @("본사확정·전표", "058S 구독 대리점 100090, 058A 일반 대리점 100273 에서 실제 수행 → 전표취소·확정취소·본사확정취소로 원복")
    Row @("선정 조건", "조치일자 2026.07.01~03 / 무상(N) / CSLTNO_ITEM=1 / 제외사유 공백 / 기사 있음 / SUBS_TYPE 공백 / 센터코드 아님 / KNVV VKORG=ASKR / FEEH1 2026.07 미처리")
    Gap

    Section "1. 구독 테스트 상담건 + 기술료/출장비 테스트 값"
    Head @("역할", "상담번호", "대리점 / 센터", "조치일자", "클레임", "기사(TECHNAME)", "기술료(MAXTECHPRICE)", "출장비(TRIPPRICE)")
    Row @("X3-1", "2026063003075", "100090 (주)평택경동 / CT04", "2026.07.01", "A", "0000130272 박혁", "23,000", "20,000") $IN
    Row @("X3-2", "2026063003277", "100090", "2026.07.01", "D2", "0000130272 박혁", "13,000", "15,000") $IN
    Row @("X3-3", "2026070100389", "100090", "2026.07.01", "D2", "0000135763 김창민", "37,000", "10,000") $IN
    Row @("Y3-1", "2026070101688", "100032 수도상사 / CT17", "2026.07.01", "D2", "0000133740", "11,000", "5,000") $IN
    Note "※ X3 는 기사 2명(0000130272 2건, 0000135763 1건)으로 나눠 기사별 금액행(D-06)·기사 여러 명(D-09)을 검증. 상담번호 202606 + 조치일자 7월로 SPMON 파생(R-04)도 함께 검증."

    Section "2. 기대값과 실측값"
    Head @("화면", "항목", "기대값", "실측값", "판정")
    Row @("058S 100090", "기술료 / 출장비", "73,000 / 45,000", "73,000 / 45,000", "PASS") $OK
    Row @("058S 100032", "기술료 / 출장비", "11,000 / 5,000", "11,000 / 5,000", "PASS") $OK
    Row @("060S 100090", "기사 0000130272 금액행", "36,000 / 35,000", "36,000 / 35,000", "PASS") $OK
    Row @("060S 100090", "기사 0000135763 금액행", "37,000 / 10,000", "37,000 / 10,000", "PASS") $OK
    Row @("060S 100090", "대리점 합계행", "73,000 / 45,000", "73,000 / 45,000", "PASS") $OK
    Row @("0101 (058S 지급대상)", "2026063003075", "23,000 / 20,000", "230.00 / 200.00", "FAIL") $NG
    Row @("0101 (058S 지급대상)", "2026063003277", "13,000 / 15,000", "130.00 / 150.00", "FAIL") $NG
    Row @("0101 (058S 지급대상)", "2026070100389", "37,000 / 10,000", "370.00 / 100.00", "FAIL") $NG
    Gap

    Section "3. 확정 데이터(ZCRM_CHRG_CNFRM) 등록·취소 순서"
    Head @("순서", "케이스", "처리구분", "상담번호 : 확정일", "RFC 결과", "그 시점 058S 100090 CFDATE1")
    Row @("1", "R-05/06/07/11/12", "공백/X/S", "오류 케이스 (DB 미반영)", "E (R-05 는 SE37 차단)", "공백 (0/3) = T-04")
    Row @("2", "R-01", "S", "2026063003075:08.05 / 2026063003277:08.06 / 2026070101688:08.04", "S, 3건 저장", "공백 (2/3) = T-05, Y3 2026.08.04 = T-08")
    Row @("3", "R-08 / R-09", "S", "빈 상담번호·9999999999999 + 2026063003277", "E (부분성공)", "변화 없음")
    Row @("4", "R-02", "S", "2026063003075:08.03 (변경)", "S, 생성이력 유지", "공백 (2/3)")
    Row @("5", "T-14 / T-16", "-", "058S 본사확정 → 전표생성 시도", "차단 '대리점 0000100090 의 확정일자가 존재하지 않습니다.'", "공백 (2/3)")
    Row @("6", "R-14", "S", "2026070100389:08.09 → 08.07 (중복 2줄)", "S, 08.07 1건", "2026.08.07 (3/3) = T-06")
    Row @("7", "T-15 / T-17", "-", "058S 전표생성 → 전표취소", "전표 3100069913 생성 후 취소", "2026.08.07")
    Row @("8", "T-07", "C", "2026070100389", "S, 1건 삭제", "공백 (2/3)")
    Row @("9", "R-13", "C", "2026070100389 (재취소)", "E '취소할 확정 데이터가 없습니다'", "변화 없음")
    Row @("10", "R-03", "C", "2026063003075 / 2026063003277", "S, 2건 삭제", "-")
    Row @("11", "R-15", "C", "2026070101688 (중복 2줄)", "S, 1건 삭제 → 테이블 0건", "-") $OK
    Gap

    Section "4. 본사확정·전표 처리 이력 (ZCRM_ASSPTFEEH1 2026.07)"
    Head @("대리점", "화면", "처리", "결과", "원복")
    Row @("100090 (구독)", "058S", "본사확정 → 전표생성 → 전표취소 → 본사확정취소", "FEEH1 15,000 등록 / 전표 3100069913 생성 후 취소", "FEEH1 행 삭제") $OK
    Row @("100273 (일반)", "058A", "본사확정 → 대리점확정 → 전표생성 → 전표취소 → 확정취소 → 본사확정취소", "FEEH1 135,000 / 확정일 2026.09.13 / 전표 3100069914 생성 후 취소", "FEEH1 행 삭제") $OK
    Gap

    Section "5. 화면 금액 (기간월 2026.07)"
    Head @("T-code", "대리점", "센터", "세팅 전", "세팅 후", "구독 본사확정 중", "원복 후")
    Row @("ZCRM058A", "100090", "CT04", "195,000", "180,000", "15,000 표시 (결함 G-01)", "195,000")
    Row @("ZCRM058S", "100090", "CT04", "0행", "15,000", "15,000", "0행")
    Row @("ZCRM058A", "100032", "CT17", "30,000", "미측정", "-", "30,000")
    Row @("ZCRM058S", "100032", "CT17", "0행", "5,000", "-", "-")
    Row @("ZCRM058A", "100273", "CT23", "135,000", "135,000 (일반)", "-", "135,000")
    Gap

    Section "6. 원복 확인"
    Head @("대상", "결과")
    Row @("ZCRM_CHRG_CNFRM", "0건") $OK
    Row @("ZCRM_ASSPTFEEH1 2026.07", "0건") $OK
    Row @("ZCRM_CICCOST_M 테스트 행", "4건 삭제 (조회 시 행 없음)") $OK
    Row @("ZCRM_CICSVC_NEW-SUBS_TYPE", "4건 공백 복원") $OK
    Row @("ZCRM058A", "100090 195,000 / 100032 30,000 / 100273 135,000, 모두 LED 빨강") $OK

    $w = @(16, 30, 30, 26, 22, 30, 20, 18)
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
